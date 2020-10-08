// <copyright file="UserStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.NewHireOnboarding.Providers
{
    using System;
    using System.Collections.Generic;
    using System.Net;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.NewHireOnboarding.Models;
    using Microsoft.Teams.Apps.NewHireOnboarding.Models.Configuration;
    using Microsoft.Teams.Apps.NewHireOnboarding.Models.EntityModels;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Implements the methods that are defined in <see cref="IUserStorageProvider"/>.
    /// Implements storage provider which helps to storage user information in Azure Table Storage.
    /// </summary>
    public class UserStorageProvider : BaseStorageProvider, IUserStorageProvider
    {
        private const string UserConfigurationTable = "UserConfiguration";

        /// <summary>
        /// Initializes a new instance of the <see cref="UserStorageProvider"/> class.
        /// </summary>
        /// <param name="options">A set of key/value application configuration properties for Azure Table Storage.</param>
        /// <param name="logger">Logger implementation to send logs to the logger service.</param>
        public UserStorageProvider(
            IOptions<StorageSettings> options,
            ILogger<UserStorageProvider> logger)
            : base(options?.Value.ConnectionString, UserConfigurationTable, logger)
        {
        }

        /// <summary>
        /// Store or update user detail in Azure Table Storage.
        /// </summary>
        /// <param name="userEntity">Represents user entity used for storage and retrieval.</param>
        /// <returns><see cref="Task"/> that represents user entity is saved or updated.</returns>
        public async Task<bool> StoreOrUpdateUserDetailAsync(UserEntity userEntity)
        {
            userEntity = userEntity ?? throw new ArgumentNullException(nameof(userEntity));

            if (string.IsNullOrWhiteSpace(userEntity.AadObjectId)
                || string.IsNullOrWhiteSpace(userEntity.ConversationId)
                || string.IsNullOrWhiteSpace(userEntity.ServiceUrl))
            {
                throw new ArgumentNullException(nameof(userEntity));
            }

            var result = await this.InsertOrReplaceUserAsync(userEntity);

            if (result == null)
            {
                return false;
            }
            else
            {
                return result.HttpStatusCode == (int)HttpStatusCode.NoContent;
            }
        }

        /// <summary>
        /// Get already saved user entity from Azure Table Storage.
        /// </summary>
        /// <param name="userAadObjectId">Azure Active Directory object id of user.</param>
        /// <returns><see cref="Task"/>Returns user entity.</returns>
        public async Task<UserEntity> GetUserDetailAsync(string userAadObjectId)
        {
            if (string.IsNullOrWhiteSpace(userAadObjectId))
            {
                return null;
            }

            await this.EnsureInitializedAsync();

            var operation = TableOperation.Retrieve<UserEntity>(userAadObjectId, userAadObjectId);
            var data = await this.CloudTable.ExecuteAsync(operation);

            return data.Result as UserEntity;
        }

        /// <summary>
        /// Get all user details based on role.
        /// </summary>
        /// <param name="userRole">User role like 0:New Hire, 1:Hiring Manager.</param>
        /// <returns>List of users details based on role.</returns>
        public async Task<IEnumerable<UserEntity>> GetAllUsersAsync(UserRole userRole)
        {
            await this.EnsureInitializedAsync();

            var users = new List<UserEntity>();
            string userRoleCondition = TableQuery.GenerateFilterConditionForInt("UserRole", QueryComparisons.Equal, (int)userRole);
            TableQuery<UserEntity> query = new TableQuery<UserEntity>().Where(userRoleCondition);
            TableContinuationToken tableContinuationToken = null;

            do
            {
                var queryResponse = await this.CloudTable.ExecuteQuerySegmentedAsync(query, tableContinuationToken);
                tableContinuationToken = queryResponse.ContinuationToken;
                users.AddRange(queryResponse?.Results);
            }
            while (tableContinuationToken != null);

            return users;
        }

        /// <summary>
        /// Get all users who opted for pair-up meeting.
        /// </summary>
        /// <returns>List of users details.</returns>
        public async Task<IEnumerable<UserEntity>> GetUsersOptedForPairUpMeetingAsync()
        {
            await this.EnsureInitializedAsync();

            var users = new List<UserEntity>();
            string optedInCondition = TableQuery.GenerateFilterConditionForBool("OptedIn", QueryComparisons.Equal, true);
            TableQuery<UserEntity> query = new TableQuery<UserEntity>().Where(optedInCondition);
            TableContinuationToken tableContinuationToken = null;

            do
            {
                var queryResponse = await this.CloudTable.ExecuteQuerySegmentedAsync(query, tableContinuationToken);
                tableContinuationToken = queryResponse.ContinuationToken;
                users.AddRange(queryResponse?.Results);
            }
            while (tableContinuationToken != null);

            return users;
        }

        /// <summary>
        /// Stores or update user details data in Azure Table Storage.
        /// </summary>
        /// <param name="entity">Holds user detail entity data.</param>
        /// <returns>A task that represents user entity data is saved or updated.</returns>
        private async Task<TableResult> InsertOrReplaceUserAsync(UserEntity entity)
        {
            await this.EnsureInitializedAsync();

            TableOperation addOrUpdateOperation = TableOperation.InsertOrReplace(entity);

            return await this.CloudTable.ExecuteAsync(addOrUpdateOperation);
        }
    }
}

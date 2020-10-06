// <copyright file="ConfigurationData.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.NewHireOnboarding.Tests.TestData
{
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.NewHireOnboarding.Models.Configuration;

    /// <summary>
    /// Configuration data settings class.
    /// </summary>
    public static class ConfigurationData
    {
        public static readonly IOptions<BotOptions> botOptions = Options.Create(new BotOptions()
        {
            MicrosoftAppId = "",
            MicrosoftAppPassword = "",
            TenantId = "12345",
            AppBaseUri = "https://at-prod-neo.azurewebsites.net",
            HumanResourceTeamId = "12345"
        });

        public static readonly IOptions<SharePointSettings> sharePointOptions = Options.Create(new SharePointSettings()
        {
            CompleteLearningPlanUrl = "https://test.com",
            ShareFeedbackFormUrl = "https://test.com",
        });

        public static readonly IOptions<AadSecurityGroupSettings> aadSecurityGroupSettings = Options.Create(new AadSecurityGroupSettings()
        {
            Id = "12345",
        });

        public static readonly IOptions<TokenSettings> tokenOptions = Options.Create(new TokenSettings()
        {
            ConnectionName = "Storage connecton string"
        });

        public static readonly IOptions<StorageSettings> storageOptions = Options.Create(new StorageSettings()
        {
            ConnectionString = "DefaultEndpointsProtocol=https;AccountName=xxxx;AccountKey=bRWkXI8TJ3MgOWqkDwvjuUfEEeYwXBcpRxO3j8TNzsL4VDIDV10cZzsC7ZK4Zvs0+p3+8PwwJH1/hWOi9DdpVQ==;EndpointSuffix=core.windows.net",
        });

        /// <summary>
        /// Azure Active Directory id of team.
        /// </summary>
        public static readonly string TeamId = "5d49fe04-270d-4d34-a4d0-0044a9a08888";
    }
}
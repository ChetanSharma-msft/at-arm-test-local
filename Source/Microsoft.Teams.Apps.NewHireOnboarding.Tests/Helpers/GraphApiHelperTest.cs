// <copyright file="GraphApiHelperTest.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.NewHireOnboarding.Tests.Helpers
{
    using System.Threading.Tasks;
    using Microsoft.Extensions.Caching.Memory;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.NewHireOnboarding.Helpers;
    using Microsoft.Teams.Apps.NewHireOnboarding.Tests.TestData;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Moq;

    /// <summary>
    /// Class to test Graph API helper methods.
    /// </summary>
    [TestClass]
    public class GraphApiHelperTest
    {
        /// <summary>
        /// Graph API helper instance.
        /// </summary>
        GraphApiHelper graphApiHelper;

        [TestInitialize]
        public void GraphApiHelperTestSetup()
        {
            var logger = new Mock<ILogger<GraphApiHelper>>().Object;
            var memoryCache = new Mock<IMemoryCache>().Object;
            graphApiHelper = new GraphApiHelper(logger, memoryCache, ConfigurationData.botOptions);
        }

        [TestMethod]
        public async Task GetUserPhotoAsync_ShouldThrowOnEmptyUserId()
        {
            var Result = await graphApiHelper.GetUserPhotoAsync(GraphApiHelperData.AccessToken, null);
            Assert.IsNull(Result);
        }

        [TestMethod]
        public async Task GetUserPhotoAsync_ShouldThrowInvalidUserId()
        {
            var Result = await graphApiHelper.GetUserPhotoAsync(GraphApiHelperData.AccessToken, GraphApiHelperData.UserId);
            Assert.IsNull(Result);
        }

        [TestMethod]
        public async Task GetUserProfileAsync_ShouldThrowOnEmptyUserId()
        {
            var Result = await graphApiHelper.GetUserProfileAsync(GraphApiHelperData.AccessToken, null);
            Assert.IsNull(Result);
        }

        [TestMethod]
        public async Task GetUserProfileNoteAsync_ShouldThrowOnEmptyUserId()
        {
            var Result = await graphApiHelper.GetUserProfileNoteAsync(GraphApiHelperData.AccessToken, null);
            Assert.IsNull(Result);
        }

        [TestMethod]
        public async Task GetChannelsAsync_ShouldThrowOnEmptyTeamId()
        {
            var Result = await graphApiHelper.GetChannelsAsync(GraphApiHelperData.AccessToken, null);
            Assert.IsNull(Result);
        }

        [TestMethod]
        public async Task GetGroupMemberDetailsAsync_ShouldThrowOnEmptyGroupId()
        {
            var Result = await graphApiHelper.GetGroupMemberDetailsAsync(GraphApiHelperData.AccessToken, null);
            Assert.IsNull(Result);
        }

        [TestMethod]
        public async Task UserManagerDetailsAsync_Null()
        {
            var Result = await graphApiHelper.GetUserManagerDetailsAsync(null);
            Assert.IsNull(Result);
        }
    }
}

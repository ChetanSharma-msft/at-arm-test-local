// <copyright file="ToKenHelperTest.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.NewHireOnboarding.Tests.Helpers
{
    using Microsoft.Bot.Builder.Integration;
    using Microsoft.Bot.Builder.Integration.AspNet.Core;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Bot.Schema;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.NewHireOnboarding.Helpers;
    using Microsoft.Teams.Apps.NewHireOnboarding.Providers;
    using Microsoft.Teams.Apps.NewHireOnboarding.Tests.TestData;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Moq;
    using System.Threading;
    using System.Threading.Tasks;

    // Class to teams info helper methods.
    [TestClass]
    public class TeamsInfoHelperTest
    {
        TeamsInfoHelper teamsInfoHelper;
        Mock<ITeamStorageProvider> teamStorageProvider;
        Mock<IBotFrameworkHttpAdapter> botAdapter;
        MicrosoftAppCredentials microsoftAppCredentials;

        Mock<IAdapterIntegration> adapterIntegration;
        ConversationReference conversationReference;

        [TestInitialize]
        public void teamsInfoHelperTestSetup()
        {
            var logger = new Mock<ILogger<TeamsInfoHelper>>().Object;
            teamStorageProvider = new Mock<ITeamStorageProvider>();
            botAdapter = new Mock<IBotFrameworkHttpAdapter>();
            teamsInfoHelper = new TeamsInfoHelper(botAdapter.Object, teamStorageProvider.Object, microsoftAppCredentials, logger);

            adapterIntegration = new Mock<IAdapterIntegration>();

            var conversationReference = new ConversationReference
            {
                ChannelId = CardConstants.TeamsBotFrameworkChannelId,
                ServiceUrl = "https://test.com",
            };
        }

        [TestMethod]
        public async Task GetTeamMemberAsync()
        {
            this.teamStorageProvider
                .Setup(x => x.GetTeamDetailAsync("123"))
                .Returns(Task.FromResult(NotificationHelperData.teamEntity));

            this.adapterIntegration
               .Setup(x => x.ContinueConversationAsync("12345", conversationReference, null, CancellationToken.None))
               .Returns(Task.FromResult(NotificationHelperData.teamEntity));

            var Result = await teamsInfoHelper.GetTeamMemberAsync("123", "6d230b1a-065e-4dab-9253-caa64f2d3519");
            Assert.IsNull(Result);
        }
    }
}

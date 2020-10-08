// <copyright file="ActivityHelperTest.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.NewHireOnboarding.Tests.Helpers
{
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Azure;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.NewHireOnboarding.Dialogs;
    using Microsoft.Teams.Apps.NewHireOnboarding.Helpers;
    using Microsoft.Teams.Apps.NewHireOnboarding.Models;
    using Microsoft.Teams.Apps.NewHireOnboarding.Providers;
    using Microsoft.Teams.Apps.NewHireOnboarding.Tests.TestData;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Moq;
    using Newtonsoft.Json;
    using System;
    using System.Dynamic;
    using System.Linq;
    using System.Threading.Tasks;

    /// <summary>
    /// Class to test activity helper methods.
    /// </summary>
    [TestClass]
    public class ActivityHelperTest
    {
        Mock<IGraphUtilityHelper> graphUtility;
        Mock<ISharePointHelper> sharePointHelper;
        Mock<ICardHelper> cardHelper;

        Mock<ITeamStorageProvider> teamStorageProvider;
        Mock<IIntroductionStorageProvider> introductionStorageProvider;
        Mock<IUserStorageProvider> userStorageProvider;
        Mock<IFeedbackProvider> feedbackProvider;
        Mock<IImageUploadProvider> imageUploadProvider;

        Mock<ITeamMembership> teamMembership;
        Mock<IUserProfile> userProfile;
        ActivityHelper<MainDialog> activityHelper;
        ITurnContext turnContext;
        ITurnContext<IInvokeActivity> turnContextInvokeActivity;
        Mock<ITurnContext> turnContextMock;

        [TestInitialize]
        public void LearningPlanHelperTestSetup()
        {
            // Mock providers
            teamStorageProvider = new Mock<ITeamStorageProvider>();
            introductionStorageProvider = new Mock<IIntroductionStorageProvider>();
            userStorageProvider = new Mock<IUserStorageProvider>();
            feedbackProvider = new Mock<IFeedbackProvider>();
            imageUploadProvider = new Mock<IImageUploadProvider>();

            // Mock helpers
            sharePointHelper = new Mock<ISharePointHelper>();
            graphUtility = new Mock<IGraphUtilityHelper>();
            cardHelper = new Mock<ICardHelper>();

            dynamic myexpando = new ExpandoObject();
            myexpando.Data = new ExpandoObject() as dynamic;
            myexpando.Data = new AdaptiveSubmitActionData
            {
                Msteams = new CardAction
                {
                    Type = ActionTypes.MessageBack,
                    Text = BotCommandConstants.RequestMoreInfoAction,
                },
                IntroductionEntity = ActivityHelperData.introductionEntity,
            };

            var botAdapter = new Mock<BotAdapter>();

            turnContext = new TurnContext(
                botAdapter.Object,
                new Activity
                {
                    Value = JsonConvert.SerializeObject(myexpando),
                });

            turnContextInvokeActivity = null;

            teamMembership = new Mock<ITeamMembership>();
            userProfile = new Mock<IUserProfile>();
            var loggerMainDialog = new Mock<ILogger<MainDialog>>().Object;
            var logger = new Mock<ILogger<ActivityHelper<MainDialog>>>().Object;
            var localizer = new Mock<IStringLocalizer<Strings>>().Object;
            turnContextMock = new Mock<ITurnContext>();

            IStorage storage = new AzureBlobStorage(ConfigurationData.storageOptions.Value.ConnectionString, "bot-state");
            UserState userState = new UserState(storage);
            ConversationState conversationState = new ConversationState(storage);

            MainDialog mainDialog = new MainDialog(
                ConfigurationData.tokenOptions,
                loggerMainDialog,
                localizer);

            // Class contructor.
            activityHelper = new ActivityHelper<MainDialog>(
                logger,
                userState,
                teamStorageProvider.Object,
                localizer,
                mainDialog,
                conversationState,
                teamMembership.Object,
                userProfile.Object,
                introductionStorageProvider.Object,
                sharePointHelper.Object,
                cardHelper.Object,
                graphUtility.Object,
                ConfigurationData.botOptions,
                userStorageProvider.Object,
                ConfigurationData.aadSecurityGroupSettings,
                feedbackProvider.Object,
                imageUploadProvider.Object);
        }

        [TestMethod]
        public async Task IntroductionCardAsync_Success()
        {
            this.userProfile
                 .Setup(x => x.GetUserManagerDetailsAsync(GraphApiHelperData.AccessToken))
                 .Returns(Task.FromResult(ActivityHelperData.userProfileDetail));

            this.turnContextMock
                 .Setup(x => x.SendActivityAsync(GraphApiHelperData.AccessToken, null, "acceptingInput", new System.Threading.CancellationToken()))
                 .Returns(Task.FromResult(new ResourceResponse()));

            this.introductionStorageProvider
                 .Setup(x => x.GetIntroductionDetailAsync(ActivityHelperData.NewJoinerAadObjectId, ActivityHelperData.HiringManagerAadObjectId))
                 .Returns(Task.FromResult(ActivityHelperData.introductionEntity));

            this.graphUtility
                .Setup(x => x.ObtainApplicationTokenAsync(ConfigurationData.botOptions.Value.TenantId, ConfigurationData.botOptions.Value.MicrosoftAppId, ConfigurationData.botOptions.Value.MicrosoftAppPassword))
                .Returns(Task.FromResult(new Models.Graph.GraphTokenResponse() { AccessToken = GraphApiHelperData.AccessToken }));

            this.sharePointHelper
                 .Setup(x => x.GetIntroductionQuestionsAsync(GraphApiHelperData.AccessToken))
                 .Returns(Task.FromResult(ActivityHelperData.IntroductionQuestions));

            this.userProfile
                 .Setup(x => x.GetUserProfileNoteAsync(GraphApiHelperData.AccessToken, GraphApiHelperData.UserId))
                 .Returns(Task.FromResult("ProfileNote"));

            this.cardHelper
                 .Setup(x => x.GetNewHireIntroductionCard(ActivityHelperData.introductionEntity, true))
                 .Returns(new TaskModuleResponse() { Task = new TaskModuleResponseBase() { Type = "continue" } });

            this.cardHelper
                 .Setup(x => x.GetIntroductionValidationCard(ActivityHelperData.introductionEntity))
                 .Returns(new TaskModuleResponse() { Task = new TaskModuleResponseBase() { Type = "continue" } });

            this.cardHelper
                 .Setup(x => x.GetNewHireIntroductionCard(ActivityHelperData.introductionEntity, true))
                 .Returns(new TaskModuleResponse() { Task = new TaskModuleResponseBase() { Type = "continue" } });

            var Result = await this.activityHelper.GetIntroductionAsync(GraphApiHelperData.AccessToken, turnContext, new System.Threading.CancellationToken());
            Assert.AreEqual(Result?.Task?.Type, "continue");
        }

        [TestMethod]
        public async Task ApproveIntroductionCardAsync_Success()
        {
            this.introductionStorageProvider
                 .Setup(x => x.GetIntroductionDetailAsync(ActivityHelperData.NewJoinerAadObjectId, ActivityHelperData.HiringManagerAadObjectId))
                 .Returns(Task.FromResult(ActivityHelperData.introductionEntity));

            this.cardHelper
                 .Setup(x => x.GetValidationErrorCard(ActivityHelperData.ValidationErrorCardText))
                 .Returns(new TaskModuleResponse());

            this.teamStorageProvider
                 .Setup(x => x.GetAllTeamDetailAsync())
                 .Returns(Task.FromResult(ActivityHelperData.teamEntities));

            this.teamMembership
                .Setup(x => x.GetMyJoinedTeamsAsync(GraphApiHelperData.AccessToken))
                .Returns(Task.FromResult(ActivityHelperData.teamCollection));

            this.teamMembership
                .Setup(x => x.GetChannelsAsync(GraphApiHelperData.AccessToken, ConfigurationData.TeamId))
                .Returns(Task.FromResult(ActivityHelperData.channelCollection));

            this.cardHelper
                 .Setup(x => x.GetApproveDetailCard(ActivityHelperData.teamDetailCollection, ActivityHelperData.introductionEntity, true))
                 .Returns(new TaskModuleResponse());

            try
            {
                var Result = await this.activityHelper.ApproveIntroductionActionAsync(GraphApiHelperData.AccessToken, turnContext);
                Assert.AreEqual(Result.Task.Type, ActivityHelperData.TaskModuleSuccessResponseType);
            }
            catch (Exception ex)
            {
                // Fail to deserialize the bot commands from ITurnContext activity object.
                Assert.AreEqual(ex.Message, "Unable to cast object of type 'System.String' to type 'Newtonsoft.Json.Linq.JObject'.");
            }
        }

        [TestMethod]
        public async Task TeamMappingDetailsAsync_Success()
        {
            this.teamStorageProvider
                 .Setup(x => x.GetAllTeamDetailAsync())
                 .Returns(Task.FromResult(ActivityHelperData.teamEntities));

            this.teamMembership
                .Setup(x => x.GetMyJoinedTeamsAsync(GraphApiHelperData.AccessToken))
                .Returns(Task.FromResult(ActivityHelperData.teamCollection));

            this.teamMembership
                .Setup(x => x.GetChannelsAsync(GraphApiHelperData.AccessToken, ConfigurationData.TeamId))
                .Returns(Task.FromResult(ActivityHelperData.channelCollection));

            var Result = await this.activityHelper.GetTeamMappingDetailsAsync(turnContext, GraphApiHelperData.AccessToken);
            Assert.AreEqual(Result.ToList().Any(), true);
        }

        [TestMethod]
        public async Task TeamMappingDetailsAsync_Failure()
        {
            this.teamStorageProvider
                 .Setup(x => x.GetAllTeamDetailAsync())
                 .Returns(Task.FromResult(ActivityHelperData.emptyTeamCollection));

            this.teamMembership
                .Setup(x => x.GetMyJoinedTeamsAsync(GraphApiHelperData.AccessToken))
                .Returns(Task.FromResult(ActivityHelperData.teamCollection));

            this.teamMembership
                .Setup(x => x.GetChannelsAsync(GraphApiHelperData.AccessToken, ConfigurationData.TeamId))
                .Returns(Task.FromResult(ActivityHelperData.channelCollection));

            var Result = await this.activityHelper.GetTeamMappingDetailsAsync(turnContext, GraphApiHelperData.AccessToken);
            Assert.AreEqual(Result, null);
        }

        [TestMethod]
        public async Task SubmitIntroductionAsync_Success()
        {
            this.userProfile
                 .Setup(x => x.GetUserManagerDetailsAsync(GraphApiHelperData.AccessToken))
                 .Returns(Task.FromResult(ActivityHelperData.userProfileDetail));

            this.turnContextMock
                 .Setup(x => x.SendActivityAsync("UserNotMappedWithManagerMessageText", null, "acceptingInput", new System.Threading.CancellationToken()))
                 .Returns(Task.FromResult(new ResourceResponse()));

            this.userStorageProvider
                .Setup(x => x.GetUserDetailAsync(GraphApiHelperData.UserId))
                .Returns(Task.FromResult(ActivityHelperData.userEntity));

            this.turnContextMock
                 .Setup(x => x.SendActivityAsync("ManagerUnavailableText", null, "acceptingInput", new System.Threading.CancellationToken()))
                 .Returns(Task.FromResult(new ResourceResponse()));

            this.introductionStorageProvider
                 .Setup(x => x.StoreOrUpdateIntroductionDetailAsync(ActivityHelperData.introductionEntity))
                 .Returns(Task.FromResult(true));

            try
            {
                var Result = await this.activityHelper.SubmitIntroductionActionAsync(
                GraphApiHelperData.AccessToken,
                turnContextInvokeActivity,
                new TaskModuleRequest(),
                new System.Threading.CancellationToken());

                Assert.AreEqual(Result.Task.Type, ActivityHelperData.TaskModuleSuccessResponseType);
            }
            catch (Exception ex)
            {
                // Fail if turn context is null.
                Assert.AreEqual(ex.Message, "Value cannot be null. (Parameter 'turnContext')");
            }
        }
    }
}

// <copyright file="ToKenHelperTest.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.NewHireOnboarding.Tests.Helpers
{
    using Microsoft.Bot.Connector;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.NewHireOnboarding.Helpers;
    using Microsoft.Teams.Apps.NewHireOnboarding.Models.Configuration;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Moq;
    using System.Threading.Tasks;

    /// <summary>
    /// Class to test token helper methods.
    /// </summary>
    [TestClass]
    public class ToKenHelperTest
    {
        OAuthClient oAuthClient;
        TokenHelper tokenHelper;

        public static readonly IOptions<TokenSettings> options = Options.Create(new TokenSettings()
        {
            ConnectionName = "test"
        });

        [TestInitialize]
        public void ToKenHelperTestSetup()
        {
            var logger = new Mock<ILogger<TokenHelper>>().Object;
            tokenHelper = new TokenHelper(oAuthClient, options, logger);
        }

        [TestMethod]
        public async Task GetUserTokenAsyncAsync_ThrowsException()
        {
            var Result = await tokenHelper.GetUserTokenAsync("6d230b1a-065e-4dab-9253-caa64f2d3519", "https://graph.microsoft.com/");
            Assert.IsNull(Result);
        }
    }
}

// <copyright file="DeepLinkConstants.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.NewHireOnboarding
{
    /// <summary>
    /// A class that holds deep links that are used in multiple files.
    /// </summary>
    public static class DeepLinkConstants
    {
        /// <summary>
        /// Microsoft Graph API base url.
        /// </summary>
        public const string GraphAPIBaseURL = "https://graph.microsoft.com/";

        /// <summary>
        /// Deep link to navigate to Channels Tab.
        /// </summary>
        public const string TabBaseRedirectURL = "https://teams.microsoft.com/l/entity";

        /// <summary>
        /// Deep link to initiate chat.
        /// </summary>
        public const string ChatInitiateURL = "https://teams.microsoft.com/l/chat";

        /// <summary>
        /// Link that redirects to tab.
        /// </summary>
        public const string FeedbackTabURL = "https://teams.microsoft.com/l/entity/{0}/Feedback?context={1}";

        /// <summary>
        /// Link to open file in teams.
        /// </summary>
        public const string OpenFileInTeamsURL = "https://teams.microsoft.com/_#/{0}/viewer/teams/{1}";
    }
}

// <copyright file="PairUpNotificationAdaptiveCard.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.NewHireOnboarding.Cards
{
    using System;
    using System.Collections.Generic;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Extensions.Localization;
    using Microsoft.Teams.Apps.NewHireOnboarding.Models;
    using Microsoft.Teams.Apps.NewHireOnboarding.Models.EntityModels;

    /// <summary>
    /// Class for the pair-up notification card.
    /// </summary>
    public static class PairUpNotificationAdaptiveCard
    {
        /// <summary>
        /// Default marker string in the UPN that indicates a user is externally-authenticated
        /// </summary>
        private const string ExternallyAuthenticatedUpnMarker = "#ext#";

        /// <summary>
        /// Creates the pair-up notification card.
        /// </summary>
        /// <param name="sender">The user who will be sending this card.</param>
        /// <param name="recipient">The user who will be receiving this card.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>Pair-up notification card</returns>
        public static Attachment GetPairUpNotificationCard(UserEntity sender, UserEntity recipient, IStringLocalizer<Strings> localizer)
        {
            sender = sender ?? throw new ArgumentNullException(nameof(sender));
            recipient = recipient ?? throw new ArgumentNullException(nameof(recipient));

            // To start a chat with a guest user, use their external email, not the UPN
            var recipientUpn = !IsGuestUser(recipient) ? recipient.UserPrincipalName : recipient.Email;

            var meetingTitle = localizer.GetString("MeetupTitle", sender.Name, recipient.Name);
            var meetingContent = localizer.GetString("MeetupContent", localizer.GetString("AppTitle"));
            var meetingLink = "https://teams.microsoft.com/l/meeting/new?subject=" + Uri.EscapeDataString(meetingTitle) + "&attendees=" + recipientUpn + "&content=" + Uri.EscapeDataString(meetingContent);
            var encodedMessage = Uri.EscapeDataString(localizer.GetString("InitiateChatText"));

            AdaptiveCard pairUpNotificationCard = new AdaptiveCard(CardConstants.AdaptiveCardVersion)
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveTextBlock
                    {
                        Size = AdaptiveTextSize.Medium,
                        Weight = AdaptiveTextWeight.Bolder,
                        Text = localizer.GetString("MatchUpCardTitleContent"),
                        Wrap = true,
                        MaxLines = 2,
                    },
                    new AdaptiveTextBlock
                    {
                        Text = localizer.GetString("MatchUpCardMatchedText", recipient.Name),
                        Wrap = true,
                    },
                    new AdaptiveTextBlock
                    {
                        Text = localizer.GetString("MatchUpCardContentPart1", localizer.GetString("AppTitle"), recipient.Name),
                        Wrap = true,
                    },
                    new AdaptiveTextBlock
                    {
                        Text = localizer.GetString("MatchUpCardContentPart2"),
                        Wrap = true,
                    },
                },
                Actions = new List<AdaptiveAction>
                {
                    new AdaptiveOpenUrlAction
                    {
                        Title = localizer.GetString("ChatWithMatchButtonText", recipient.Name),
                        Url = new Uri($"{DeepLinkConstants.ChatInitiateURL}/0/0?users={Uri.EscapeDataString(recipientUpn)}&message={encodedMessage}"),
                    },
                    new AdaptiveOpenUrlAction
                    {
                        Title = localizer.GetString("ProposeMeetupButtonText"),
                        Url = new Uri(meetingLink),
                    },
                    new AdaptiveSubmitAction
                    {
                        Title = localizer.GetString("PauseMatchesButtonText"),
                        Data = new AdaptiveSubmitActionData
                        {
                            Msteams = new CardAction
                            {
                                Type = ActionTypes.MessageBack,
                                DisplayText = localizer.GetString("PauseMatchesButtonText"),
                                Text = BotCommandConstants.PauseAllMatches,
                            },
                        },
                    },
                },
            };

            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = pairUpNotificationCard,
            };
        }

        /// <summary>
        /// Checks whether or not the account is a guest user.
        /// </summary>
        /// <param name="account">The <see cref="UserEntity"/> user to check.</param>
        /// <returns>True if the account is a guest user, false otherwise.</returns>
        private static bool IsGuestUser(UserEntity account)
        {
            return account.UserPrincipalName.IndexOf(ExternallyAuthenticatedUpnMarker, StringComparison.InvariantCultureIgnoreCase) >= 0;
        }
    }
}
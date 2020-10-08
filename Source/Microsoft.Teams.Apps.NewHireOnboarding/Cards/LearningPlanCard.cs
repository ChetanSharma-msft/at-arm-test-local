﻿// <copyright file="LearningPlanCard.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.NewHireOnboarding.Cards
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Extensions.Localization;
    using Microsoft.Teams.Apps.NewHireOnboarding.Models;
    using Microsoft.Teams.Apps.NewHireOnboarding.Models.SharePoint;

    /// <summary>
    /// Class that helps to return learning card for new hire as attachment.
    /// </summary>
    public static class LearningPlanCard
    {
        /// <summary>
        /// Represents image height in pixel.
        /// </summary>
        private const int ImageHeight = 132;

        /// <summary>
        /// Represents image width in pixel.
        /// </summary>
        private const int ImageWidth = 500;

        /// <summary>
        /// Get learning card attachment for new hire to show on Microsoft Teams personal scope.
        /// </summary>
        /// <param name="localizer">The current culture's string localizer.</param>
        /// <param name="appBasePath">Application base uri to create image path.</param>
        /// <param name="learningPlan">Learning plan details.</param>
        /// <returns>New hire learning card attachment.</returns>
        public static Attachment GetNewHireLearningCard(
            IStringLocalizer<Strings> localizer,
            string appBasePath,
            LearningPlanListItemField learningPlan)
        {
            learningPlan = learningPlan ?? throw new ArgumentNullException(nameof(learningPlan));

            AdaptiveCard card = new AdaptiveCard(new AdaptiveSchemaVersion(CardConstants.AdaptiveCardVersion))
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Stretch,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Weight = AdaptiveTextWeight.Bolder,
                                        Text = learningPlan.Topic,
                                        Wrap = true,
                                        Size = AdaptiveTextSize.ExtraLarge,
                                    },
                                    new AdaptiveTextBlock
                                    {
                                        Spacing = AdaptiveSpacing.Small,
                                        Size = AdaptiveTextSize.Small,
                                        Text = learningPlan.TaskName,
                                        Color = AdaptiveTextColor.Accent,
                                        Wrap = true,
                                    },
                                    new AdaptiveImage
                                    {
                                        Url = string.IsNullOrEmpty(learningPlan.TaskImage.Url)
                                        ? new Uri($"{appBasePath}/Artifacts/learningPlan.png")
                                        : new Uri(learningPlan.TaskImage.Url),
                                        AltText = learningPlan.Notes,
                                        PixelHeight = ImageHeight,
                                        PixelWidth = ImageWidth,
                                    },
                                    new AdaptiveTextBlock
                                    {
                                        Spacing = AdaptiveSpacing.Medium,
                                        Text = learningPlan?.Notes,
                                        Wrap = true,
                                    },
                                },
                            },
                        },
                    },
                },
                Actions = new List<AdaptiveAction>(),
            };

            var actionDataUrl = learningPlan?.Link?.Url?.Replace("/", "~2F", StringComparison.InvariantCultureIgnoreCase);
            var filExtension = string.Empty;

            if (!string.IsNullOrEmpty(learningPlan?.Link?.Url))
            {
                filExtension = GetFileExtensionFromUrl(learningPlan?.Link?.Url)?.Split(".")[1];

                card.Actions.Add(
                    new AdaptiveOpenUrlAction
                    {
                        Title = localizer.GetString("ViewLearningPlanButtonText"),
                        Url = string.IsNullOrEmpty(filExtension)
                    ? new Uri(learningPlan.Link.Url)
                    : new Uri(string.Format(CultureInfo.InvariantCulture, DeepLinkConstants.OpenFileInTeamsURL, filExtension, actionDataUrl)),
                    });
            }

            card.Actions.Add(
                new AdaptiveSubmitAction
                {
                    Title = localizer.GetString("LearningPlanShareFeedbackButtonText"),
                    Data = new AdaptiveSubmitActionData
                    {
                        Msteams = new CardAction
                        {
                            Type = ActionTypes.MessageBack,
                            Text = BotCommandConstants.ShareFeedback,
                        },
                    },
                });

            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = card,
            };
        }

        /// <summary>
        /// Get file extension from content URL.
        /// </summary>
        /// <param name="url">Learning content URL.</param>
        /// <returns>File extension.</returns>
        private static string GetFileExtensionFromUrl(string url)
        {
            url = url.Split('?')[0];
            url = url.Split('/').Last();
            return url.Contains('.', StringComparison.InvariantCultureIgnoreCase) ? url.Substring(url.LastIndexOf('.')) : string.Empty;
        }
    }
}

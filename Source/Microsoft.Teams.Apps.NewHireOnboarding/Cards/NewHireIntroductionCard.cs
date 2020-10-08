// <copyright file="NewHireIntroductionCard.cs" company="Microsoft">
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
    using Newtonsoft.Json;

    /// <summary>
    /// Class that helps to return introduction card as attachment.
    /// </summary>
    public static class NewHireIntroductionCard
    {
        /// <summary>
        /// Represents image height in pixel.
        /// </summary>
        private const int ImageHeight = 363;

        /// <summary>
        /// Represents image width in pixel.
        /// </summary>
        private const int ImageWidth = 218;

        /// <summary>
        /// Get new hire introduction card attachment to show on Microsoft Teams personal scope.
        /// </summary>
        /// <param name="introductionEntity">New hire introduction details.</param>
        /// <param name="localizer">The current culture's string localizer.</param>
        /// <param name="applicationBasePath">Application base path to get the logo of the application.</param>
        /// <param name="isAllQuestionsAnswered">False when any of the question is not answered.</param>
        /// <returns>New Hire Introduction Card attachment.</returns>
        public static Attachment GetNewHireIntroductionCardAttachment(IntroductionEntity introductionEntity, IStringLocalizer<Strings> localizer, string applicationBasePath, bool isAllQuestionsAnswered = true)
        {
            introductionEntity = introductionEntity ?? throw new ArgumentNullException(nameof(introductionEntity));

            AdaptiveCard card = new AdaptiveCard(new AdaptiveSchemaVersion(CardConstants.AdaptiveCardVersion))
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveTextBlock
                    {
                        Spacing = AdaptiveSpacing.Medium,
                        Text = localizer.GetString("IntroductionText"),
                        Wrap = true,
                    },
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Auto,
                                Height = AdaptiveHeight.Stretch,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveImage
                                    {
                                        Url = new Uri($"{applicationBasePath}/Artifacts/newHireIntroduction.png"),
                                        AltText = localizer.GetString("AltTextForIntroductionCardImage"),
                                        PixelHeight = ImageHeight,
                                        PixelWidth = ImageWidth,
                                    },
                                },
                            },
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Stretch,
                                Items = GetDynamicQuestionsList(introductionEntity, localizer, isAllQuestionsAnswered),
                            },
                        },
                    },
                },

                Actions = new List<AdaptiveAction>
                {
                    new AdaptiveSubmitAction
                    {
                        Title = localizer.GetString("IntroductionSubmitButtonText"),
                        Data = new AdaptiveSubmitActionData
                        {
                            Command = BotCommandConstants.SubmitIntroductionAction,
                        },
                    },
                },
            };

            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = card,
            };
        }

        /// <summary>
        /// Get list of dynamic adaptive elements for new hire introduction card.
        /// </summary>
        /// <param name="introductionEntity">New hire introduction details.</param>
        /// <param name="localizer">The current culture's string localizer.</param>
        /// <param name="isAllQuestionsAnswered">False when any of the question is not answered.</param>
        /// <returns>List of adaptive elements.</returns>
        private static List<AdaptiveElement> GetDynamicQuestionsList(IntroductionEntity introductionEntity, IStringLocalizer<Strings> localizer, bool isAllQuestionsAnswered = true)
        {
            List<IntroductionDetail> questionAnswerList = JsonConvert.DeserializeObject<List<IntroductionDetail>>(introductionEntity.NewHireQuestionnaire);
            List<AdaptiveElement> adaptiveElements = new List<AdaptiveElement>();

            adaptiveElements.Add(
                new AdaptiveTextBlock
                {
                    Size = AdaptiveTextSize.Small,
                    Spacing = AdaptiveSpacing.Small,
                    Text = localizer.GetString("IntroductionHeaderText"),
                    Wrap = true,
                });

            adaptiveElements.Add(
                 new AdaptiveTextInput
                 {
                     Spacing = AdaptiveSpacing.Small,
                     Value = !string.IsNullOrWhiteSpace(introductionEntity.NewHireProfileNote) ? introductionEntity.NewHireProfileNote : localizer.GetString("IntroductionGreetText", introductionEntity.NewHireName),
                     Id = CardConstants.NewHireProfileNoteInputId,
                     MaxLength = 500,
                 });

            foreach (var questionAnswer in questionAnswerList)
            {
                var question = new AdaptiveTextBlock
                {
                    Size = AdaptiveTextSize.Medium,
                    Text = questionAnswer.Question,
                    Wrap = true,
                    Spacing = AdaptiveSpacing.Medium,
                };

                var answer = new AdaptiveTextInput
                {
                    Id = $"{CardConstants.QuestionId}{questionAnswerList.IndexOf(questionAnswer)}",
                    Spacing = AdaptiveSpacing.Small,
                    Value = !string.IsNullOrWhiteSpace(questionAnswer.Answer) ? questionAnswer.Answer : string.Empty,
                    MaxLength = 500,
                    Placeholder = localizer.GetString("IntroductionInputPlaceholderText"),
                };

                adaptiveElements.Add(question);
                adaptiveElements.Add(answer);
            }

            adaptiveElements.Add(
                new AdaptiveTextBlock
                {
                    Text = localizer.GetString("ValidationMessageText"),
                    Spacing = AdaptiveSpacing.Medium,
                    IsVisible = !isAllQuestionsAnswered,
                    Color = AdaptiveTextColor.Attention,
                    Size = AdaptiveTextSize.Small,
                });

            return adaptiveElements;
        }
    }
}

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.Dialogs.Choices;
using Microsoft.Bot.Schema;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using CalendarBot.Graph;
using AdaptiveCards;

namespace CalendarBot.Dialogs
{
    public class MainDialog : LogoutDialog
    {
        const string NO_PROMPT = "no-prompt";
        protected readonly ILogger _logger;
        private readonly IGraphClientService _graphClientService;

        // <ConstructorSignatureSnippet>
        public MainDialog(
            IConfiguration configuration,
            ILogger<MainDialog> logger,
            IGraphClientService graphClientService)
            : base(nameof(MainDialog), configuration["ConnectionName"])
        // </ConstructorSignatureSnippet>
        {
            _logger = logger;
            _graphClientService = graphClientService;

            // OAuthPrompt dialog handles the authentication and token
            // acquisition
            AddDialog(new OAuthPrompt(
                nameof(OAuthPrompt),
                new OAuthPromptSettings
                {
                    ConnectionName = ConnectionName,
                    Text = "Please login",
                    Title = "Login",
                    Timeout = 300000, // User has 5 minutes to login
                }));

            AddDialog(new ChoicePrompt(nameof(ChoicePrompt)));

            AddDialog(new NewEventDialog(configuration, graphClientService));

            AddDialog(new WaterfallDialog(nameof(WaterfallDialog), new WaterfallStep[]
            {
                LoginPromptStepAsync,
                ProcessLoginStepAsync,
                PromptUserStepAsync,
                CommandStepAsync,
                ProcessStepAsync,
                ReturnToPromptStepAsync
            }));

            // The initial child Dialog to run.
            InitialDialogId = nameof(WaterfallDialog);
        }

        private async Task<DialogTurnResult> LoginPromptStepAsync(
            WaterfallStepContext stepContext,
            CancellationToken cancellationToken)
        {
            // If we're going through the waterfall a second time, don't do an extra OAuthPrompt
            var options = stepContext.Options?.ToString();
            if (options == NO_PROMPT)
            {
                return await stepContext.NextAsync(cancellationToken: cancellationToken);
            }

            return await stepContext.BeginDialogAsync(nameof(OAuthPrompt), null, cancellationToken);
        }

        private async Task<DialogTurnResult> ProcessLoginStepAsync(
            WaterfallStepContext stepContext,
            CancellationToken cancellationToken)
        {
            // If we're going through the waterfall a second time, don't do an extra OAuthPrompt
            var options = stepContext.Options?.ToString();
            if (options == NO_PROMPT)
            {
                return await stepContext.NextAsync(cancellationToken: cancellationToken);
            }

            // Get the token from the previous step. If it's there, login was successful
            if (stepContext.Result != null)
            {
                var tokenResponse = stepContext.Result as TokenResponse;
                if (!string.IsNullOrEmpty(tokenResponse?.Token))
                {
                    await stepContext.Context.SendActivityAsync(
                        MessageFactory.Text("You are now logged in."), cancellationToken);
                    return await stepContext.NextAsync(null, cancellationToken);
                }
            }

            await stepContext.Context.SendActivityAsync(
                MessageFactory.Text("Login was not successful please try again."), cancellationToken);
            return await stepContext.EndDialogAsync();
        }

        private async Task<DialogTurnResult> PromptUserStepAsync(
            WaterfallStepContext stepContext,
            CancellationToken cancellationToken)
        {
            var options = new PromptOptions
            {
                Prompt = MessageFactory.Text("Please choose an option below"),
                Choices = new List<Choice> {
                    new Choice { Value = "Show token" },
                    new Choice { Value = "Show me" },
                    new Choice { Value = "Show calendar" },
                    new Choice { Value = "Add event" },
                    new Choice { Value = "Log out" },
                }
            };

            return await stepContext.PromptAsync(
                nameof(ChoicePrompt),
                options,
                cancellationToken);
        }

        private async Task<DialogTurnResult> CommandStepAsync(
            WaterfallStepContext stepContext,
            CancellationToken cancellationToken)
        {
            // Save the command the user entered so we can get it back after
            // the OAuthPrompt completes
            var foundChoice = stepContext.Result as FoundChoice;
            // Result could be a FoundChoice (if user selected a choice button)
            // or a string (if user just typed something)
            stepContext.Values["command"] = foundChoice?.Value ?? stepContext.Result;

            // There is no reason to store the token locally in the bot because we can always just call
            // the OAuth prompt to get the token or get a new token if needed. The prompt completes silently
            // if the user is already signed in.
            return await stepContext.BeginDialogAsync(nameof(OAuthPrompt), null, cancellationToken);
        }

        private async Task<DialogTurnResult> ProcessStepAsync(
            WaterfallStepContext stepContext,
            CancellationToken cancellationToken)
        {
            if (stepContext.Result != null)
            {
                var tokenResponse = stepContext.Result as TokenResponse;

                // If we have the token use the user is authenticated so we may use it to make API calls.
                if (tokenResponse?.Token != null)
                {
                    var command = ((string)stepContext.Values["command"] ?? string.Empty).ToLowerInvariant();

                    if (command.StartsWith("show token"))
                    {
                        // Show the user's token - for testing and troubleshooting
                        // Generally production apps should not display access tokens
                        await stepContext.Context.SendActivityAsync(
                            MessageFactory.Text($"Your token is: {tokenResponse.Token}"),
                            cancellationToken);
                    }
                    // <ShowMeSnippet>
                    else if (command.StartsWith("show me"))
                    {
                        await DisplayLoggedInUser(tokenResponse.Token, stepContext, cancellationToken);
                    }
                    // </ShowMeSnippet>
                    // <ShowCalendarSnippet>
                    else if (command.StartsWith("show calendar"))
                    {
                        await DisplayCalendarView(tokenResponse.Token, stepContext, cancellationToken);
                    }
                    // </ShowCalendarSnippet>
                    // <AddEventSnippet>
                    else if (command.StartsWith("add event"))
                    {
                        return await stepContext.BeginDialogAsync(nameof(NewEventDialog), null, cancellationToken);
                    }
                    // </AddEventSnippet>
                    else
                    {
                        await stepContext.Context.SendActivityAsync(
                            MessageFactory.Text("I'm sorry, I didn't understand. Please try again."),
                            cancellationToken);
                    }
                }
            }
            else
            {
                await stepContext.Context.SendActivityAsync(
                    MessageFactory.Text("We couldn't log you in. Please try again later."),
                    cancellationToken);
                return await stepContext.EndDialogAsync(cancellationToken: cancellationToken);
            }

            // Go to the next step
            return await stepContext.NextAsync(cancellationToken: cancellationToken);
        }

        private async Task<DialogTurnResult> ReturnToPromptStepAsync(
            WaterfallStepContext stepContext,
            CancellationToken cancellationToken)
        {
            // Restart the dialog, but skip the initial login prompt
            return await stepContext.ReplaceDialogAsync(InitialDialogId, NO_PROMPT, cancellationToken);
        }

        // <DisplayLoggedInUserSnippet>
        private async Task DisplayLoggedInUser(
            string accessToken,
            WaterfallStepContext stepContext,
            CancellationToken cancellationToken)
        {
            var graphClient = _graphClientService
                .GetAuthenticatedGraphClient(accessToken);

            // Get the user
            // GET /me?$select=displayName,mail,userPrincipalName
            var user = await graphClient.Me
                .Request()
                .Select(u => new {
                    u.DisplayName,
                    u.Mail,
                    u.UserPrincipalName
                })
                .GetAsync();

            // Get the user's photo
            // GET /me/photos/48x48/$value
            var userPhoto = await graphClient.Me
                .Photos["48x48"]
                .Content
                .Request()
                .GetAsync();

            // Generate an Adaptive Card
            var userCard = CardHelper.GetUserCard(user, userPhoto);

            // Create an attachment message to send the card
            var userMessage = MessageFactory.Attachment(
                new Microsoft.Bot.Schema.Attachment {
                    ContentType = AdaptiveCard.ContentType,
                    Content = userCard
                });

            await stepContext.Context.SendActivityAsync(userMessage, cancellationToken);
        }
        // </DisplayLoggedInUserSnippet>

        // <DisplayCalendarViewSnippet>
        private async Task DisplayCalendarView(
            string accessToken,
            WaterfallStepContext stepContext,
            CancellationToken cancellationToken)
        {
            var graphClient = _graphClientService
                .GetAuthenticatedGraphClient(accessToken);

            // Get user's preferred time zone and format
            var user = await graphClient.Me
                .Request()
                .Select(u => new { u.MailboxSettings })
                .GetAsync();

            var dateTimeFormat =
                $"{user.MailboxSettings.DateFormat} {user.MailboxSettings.TimeFormat}";
            if (string.IsNullOrWhiteSpace(dateTimeFormat)) {
                // Default to a standard format if user's preference not set
                dateTimeFormat = "G";
            }

            var preferredTimeZone = user.MailboxSettings.TimeZone;
            var userTimeZone = TimeZoneInfo.FindSystemTimeZoneById(preferredTimeZone);

            var now = DateTime.UtcNow;
            // Calculate the end of the week (Sunday, midnight)
            int diff = 7 - (int)DateTime.Today.DayOfWeek;
            var weekEndUnspecified = DateTime.SpecifyKind(
                DateTime.Today.AddDays(diff), DateTimeKind.Unspecified);
            var endOfWeek = TimeZoneInfo.ConvertTimeToUtc(weekEndUnspecified, userTimeZone);

            // Set query parameters for the calendar view request
            var viewOptions = new List<QueryOption>
            {
                new QueryOption("startDateTime", now.ToString("o")),
                new QueryOption("endDateTime", endOfWeek.ToString("o"))
            };

            // Get events happening between right now and the end of the week
            // GET /me/calendarView?startDateTime=""&endDateTime=""
            var events = await graphClient.Me
                .CalendarView
                .Request(viewOptions)
                // Send user time zone in request so date/time in
                // response will be in preferred time zone
                .Header("Prefer", $"outlook.timezone=\"{preferredTimeZone}\"")
                // Get max 3 per request
                .Top(3)
                // Only return fields app will use
                .Select(e => new
                {
                    e.Subject,
                    e.Organizer,
                    e.Start,
                    e.End,
                    e.Location
                })
                // Order results chronologically
                .OrderBy("start/dateTime")
                .GetAsync();

            var calendarViewMessage = MessageFactory.Text("Here are your upcoming events");
            calendarViewMessage.AttachmentLayout = AttachmentLayoutTypes.List;

            foreach(var calendarEvent in events.CurrentPage)
            {
                var eventCard = CardHelper.GetEventCard(calendarEvent, dateTimeFormat);

                // Add the card to the message's attachments
                calendarViewMessage.Attachments.Add(new Microsoft.Bot.Schema.Attachment
                {
                    ContentType = AdaptiveCard.ContentType,
                    Content = eventCard
                });
            }

            await stepContext.Context.SendActivityAsync(calendarViewMessage, cancellationToken);
        }
        // </DisplayCalendarViewSnippet>
    }
}

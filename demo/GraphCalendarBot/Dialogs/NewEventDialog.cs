// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Schema;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using CalendarBot.Graph;
using Microsoft.Recognizers.Text.DataTypes.TimexExpression;
using TimexTypes = Microsoft.Recognizers.Text.DataTypes.TimexExpression.Constants.TimexTypes;

namespace CalendarBot.Dialogs
{
    public class NewEventDialog : LogoutDialog
    {
        protected readonly ILogger _logger;
        private readonly IGraphClientService _graphClientService;

        public NewEventDialog(
            IConfiguration configuration,
            IGraphClientService graphClientService)
            : base(nameof(NewEventDialog), configuration["ConnectionName"])
        {
            // <ConstructorSnippet>
            _graphClientService = graphClientService;

            // OAuthPrompt dialog handles the token
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

            AddDialog(new TextPrompt("subjectPrompt"));
            // Validator ensures that the input is a semi-colon delimited
            // list of email addresses
            AddDialog(new TextPrompt("attendeesPrompt", AttendeesPromptValidatorAsync));
            // Validator ensures that the input is a valid date and time
            AddDialog(new DateTimePrompt("startPrompt", StartPromptValidatorAsync));
            // Validator ensures that the input is a valid date and time
            // and that it is later than the start
            AddDialog(new DateTimePrompt("endPrompt", EndPromptValidatorAsync));
            AddDialog(new ConfirmPrompt(nameof(ConfirmPrompt)));

            AddDialog(new WaterfallDialog(nameof(WaterfallDialog), new WaterfallStep[]
            {
                PromptForSubjectAsync,
                PromptForAddAttendeesAsync,
                PromptForAttendeesAsync,
                PromptForStartAsync,
                PromptForEndAsync,
                ConfirmNewEventAsync,
                GetTokenAsync,
                AddEventAsync
            }));

            // The initial child Dialog to run.
            InitialDialogId = nameof(WaterfallDialog);
            // </ConstructorSnippet>
        }

        // <PromptForSubjectSnippet>
        private async Task<DialogTurnResult> PromptForSubjectAsync(
            WaterfallStepContext stepContext,
            CancellationToken cancellationToken)
        {
            return await stepContext.PromptAsync("subjectPrompt",
                new PromptOptions{
                    Prompt = MessageFactory.Text("What's the subject for your event?")
                },
                cancellationToken);
        }
        // </PromptForSubjectSnippet>

        // <PromptForAddAttendeesSnippet>
        private async Task<DialogTurnResult> PromptForAddAttendeesAsync(
            WaterfallStepContext stepContext,
            CancellationToken cancellationToken)
        {
            stepContext.Values["subject"] = (string)stepContext.Result;

            return await stepContext.PromptAsync(nameof(ConfirmPrompt),
                new PromptOptions{
                    Prompt = MessageFactory.Text("Do you want to invite other people to this event?")
                },
                cancellationToken);
        }
        // </PromptForAddAttendeesSnippet>

        // <PromptForAttendeesSnippet>
        private async Task<DialogTurnResult> PromptForAttendeesAsync(
            WaterfallStepContext stepContext,
            CancellationToken cancellationToken)
        {
            if ((bool)stepContext.Result)
            {
                // user wants to invite attendees
                // prompt for email addresses
                return await stepContext.PromptAsync("attendeesPrompt",
                new PromptOptions{
                    Prompt = MessageFactory.Text("Enter one or more email addresses of the people you want to invite. Separate multiple addresses with a semi-colon (;)."),
                    RetryPrompt = MessageFactory.Text("One or more email addresses you entered are not valid. Please try again.")
                },
                cancellationToken);
            }
            else
            {
                // Skip attendees prompt
                return await stepContext.NextAsync(null, cancellationToken);
            }
        }
        // </PromptForAttendeesSnippet>

        // <PromptForStartSnippet>
        private async Task<DialogTurnResult> PromptForStartAsync(
            WaterfallStepContext stepContext,
            CancellationToken cancellationToken)
        {
            stepContext.Values["attendees"] = (string)stepContext.Result;

            return await stepContext.PromptAsync("startPrompt",
                new PromptOptions{
                    Prompt = MessageFactory.Text("When does the event start?"),
                    RetryPrompt = MessageFactory.Text("I'm sorry, I didn't get that. Please provide both a day and a time.")
                },
                cancellationToken);
        }
        // </PromptForStartSnippet>

        // <PromptForEndSnippet>
        private async Task<DialogTurnResult> PromptForEndAsync(
            WaterfallStepContext stepContext,
            CancellationToken cancellationToken)
        {
            var dateTimes = stepContext.Result as IList<DateTimeResolution>;

            var start = GetDateTimeFromResolutions(dateTimes);

            stepContext.Values["start"] = start;

            return await stepContext.PromptAsync("endPrompt",
                new PromptOptions{
                    Prompt = MessageFactory.Text("When does the event end?"),
                    RetryPrompt = MessageFactory.Text("I'm sorry, I didn't get that. Please provide both a day and a time, and ensure that it is later than the start."),
                    Validations = start
                },
                cancellationToken);
        }
        // </PromptForEndSnippet>

        // <ConfirmNewEventSnippet>
        private async Task<DialogTurnResult> ConfirmNewEventAsync(
            WaterfallStepContext stepContext,
            CancellationToken cancellationToken)
        {
            var dateTimes = stepContext.Result as IList<DateTimeResolution>;

            var end = GetDateTimeFromResolutions(dateTimes);

            stepContext.Values["end"] = end;

            // Playback the values as we understand them
            var subject = stepContext.Values["subject"] as string;
            var attendees = stepContext.Values["attendees"] as string;
            var start = stepContext.Values["start"] as DateTime?;

            // Build a Markdown string
            var markdown = "Here's what I heard:\n\n";
            markdown += $"- **Subject:** {subject}\n";
            markdown += $"- **Attendees:** {attendees ?? "none"}\n";
            markdown += $"- **Start:** {start?.ToString()}\n";
            markdown += $"- **End:** {end.ToString()}";

            await stepContext.Context.SendActivityAsync(
                MessageFactory.Text(markdown));

            return await stepContext.PromptAsync(nameof(ConfirmPrompt),
                new PromptOptions {
                    Prompt = MessageFactory.Text("Is this correct?")
                },
                cancellationToken);
        }
        // </ConfirmNewEventSnippet>

        // <GetTokenSnippet>
        private async Task<DialogTurnResult> GetTokenAsync(
            WaterfallStepContext stepContext,
            CancellationToken cancellationToken)
        {
            if ((bool)stepContext.Result)
            {
                return await stepContext.BeginDialogAsync(nameof(OAuthPrompt), null, cancellationToken);
            }
            else
            {
                await stepContext.Context.SendActivityAsync(
                    MessageFactory.Text("Please try again."));

                return await stepContext.EndDialogAsync(cancellationToken: cancellationToken);
            }
        }
        // </GetTokenSnippet>

        // <AddEventSnippet>
        private async Task<DialogTurnResult> AddEventAsync(
            WaterfallStepContext stepContext,
            CancellationToken cancellationToken)
        {
            if (stepContext.Result != null)
            {
                var tokenResponse = stepContext.Result as TokenResponse;
                if (tokenResponse?.Token != null)
                {
                    var subject = stepContext.Values["subject"] as string;
                    var attendees = stepContext.Values["attendees"] as string;
                    var start = stepContext.Values["start"] as DateTime?;
                    var end = stepContext.Values["end"] as DateTime?;

                    // Get an authenticated Graph client using
                    // the access token
                    var graphClient = _graphClientService
                        .GetAuthenticatedGraphClient(tokenResponse?.Token);

                    try
                    {
                        // Get user's preferred time zone
                        var user = await graphClient.Me
                            .Request()
                            .Select(u => new { u.MailboxSettings })
                            .GetAsync();

                        // Initialize an Event object
                        var newEvent = new Event
                        {
                            Subject = subject,
                            Start = new DateTimeTimeZone
                            {
                                DateTime = start?.ToString("o"),
                                TimeZone = user.MailboxSettings.TimeZone
                            },
                            End = new DateTimeTimeZone
                            {
                                DateTime = end?.ToString("o"),
                                TimeZone = user.MailboxSettings.TimeZone
                            }
                        };

                        // If attendees were provided, add them
                        if (!string.IsNullOrEmpty(attendees))
                        {
                            // Initialize a list
                            var attendeeList = new List<Attendee>();

                            // Split the string into an array
                            var emails = attendees.Split(";");
                            foreach (var email in emails)
                            {
                                // Skip empty strings
                                if (!string.IsNullOrEmpty(email))
                                {
                                    // Build a new Attendee object and
                                    // add to the list
                                    attendeeList.Add(new Attendee {
                                        Type = AttendeeType.Required,
                                        EmailAddress = new EmailAddress
                                        {
                                            Address = email
                                        }
                                    });
                                }
                            }

                            newEvent.Attendees = attendeeList;
                        }

                        // Add the event
                        // POST /me/events
                        await graphClient.Me
                            .Events
                            .Request()
                            .AddAsync(newEvent);

                        await stepContext.Context.SendActivityAsync(
                            MessageFactory.Text("Event added"),
                            cancellationToken);
                    }
                    catch (ServiceException ex)
                    {
                        _logger.LogError(ex, "Could not add event");
                        await stepContext.Context.SendActivityAsync(
                            MessageFactory.Text("Something went wrong. Please try again."),
                            cancellationToken);
                    }
                    return await stepContext.EndDialogAsync(cancellationToken: cancellationToken);
                }
            }

            await stepContext.Context.SendActivityAsync(
                    MessageFactory.Text("We couldn't log you in. Please try again later."),
                    cancellationToken);
            return await stepContext.EndDialogAsync(cancellationToken: cancellationToken);
        }
        // </AddEventSnippet>

        // Generate a DateTime from the list of
        // DateTimeResolutions provided by the DateTimePrompt
        private static DateTime GetDateTimeFromResolutions(IList<DateTimeResolution> resolutions)
        {
            var timex = new TimexProperty(resolutions[0].Timex);

            // Handle the "now" case
            if (timex.Now ?? false)
            {
                return DateTime.Now;
            }

            // Otherwise generate a DateTime
            return TimexHelpers.DateFromTimex(timex);
        }

        // <AttendeesValidatorSnippet>
        private static Task<bool> AttendeesPromptValidatorAsync(
            PromptValidatorContext<string> promptContext,
            CancellationToken cancellationToken)
        {
            if (promptContext.Recognized.Succeeded)
            {
                // Check if these are emails
                var emails = promptContext.Recognized.Value.Split(";");

                foreach (var email in emails)
                {
                    // Skip empty entries
                    if (string.IsNullOrEmpty(email))
                    {
                        continue;
                    }

                    // If there's no '@' symbol it's invalid
                    if (email.IndexOf('@') <= 0)
                    {
                        return Task.FromResult(false);
                    }

                    try
                    {
                        // Let the System.Net.Mail.MailAddress class
                        // validate the rest. If invalid it will throw
                        var mailAddress = new System.Net.Mail.MailAddress(email);
                        if (mailAddress.Address != email)
                        {
                            return Task.FromResult(false);
                        }
                    }
                    catch
                    {
                        return Task.FromResult(false);
                    }
                }

                return Task.FromResult(true);
            }

            return Task.FromResult(false);
        }
        // </AttendeesValidatorSnippet>

        // <StartValidatorSnippet>
        private static bool TimexHasDateAndTime(TimexProperty timex)
        {
            return timex.Now ?? false ||
                (timex.Types.Contains(TimexTypes.DateTime) &&
                timex.Types.Contains(TimexTypes.Definite));
        }

        private static Task<bool> StartPromptValidatorAsync(
            PromptValidatorContext<IList<DateTimeResolution>> promptContext,
            CancellationToken cancellationToken)
        {
            if (promptContext.Recognized.Succeeded)
            {
                // Initialize a TimexProperty from the first
                // recognized value
                var timex = new TimexProperty(
                    promptContext.Recognized.Value[0].Timex);

                // If it has a definite date and time, it's valid
                return Task.FromResult(TimexHasDateAndTime(timex));
            }

            return Task.FromResult(false);
        }
        // </StartValidatorSnippet>

        // <EndValidatorSnippet>
        private static Task<bool> EndPromptValidatorAsync(
            PromptValidatorContext<IList<DateTimeResolution>> promptContext,
            CancellationToken cancellationToken)
        {
            if (promptContext.Recognized.Succeeded)
            {
                if (promptContext.Options.Validations is DateTime start)
                {
                    // Initialize a TimexProperty from the first
                    // recognized value
                    var timex = new TimexProperty(
                        promptContext.Recognized.Value[0].Timex);

                    // Get the DateTime from this value to compare with start
                    var end = GetDateTimeFromResolutions(promptContext.Recognized.Value);

                    // If it has a definite date and time, and
                    // the value is later than start, it's valid
                    return Task.FromResult(TimexHasDateAndTime(timex) &&
                        DateTime.Compare(start, end) < 0);
                }
            }

            return Task.FromResult(false);
        }
        // </EndValidatorSnippet>
    }
}
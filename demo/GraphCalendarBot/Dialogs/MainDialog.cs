// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.Dialogs.Choices;
using Microsoft.Bot.Schema;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace CalendarBot.Dialogs
{
    public class MainDialog : LogoutDialog
    {
        const string NO_PROMPT = "no-prompt";
        protected readonly ILogger _logger;

        public MainDialog(IConfiguration configuration, ILogger<MainDialog> logger)
            : base(nameof(MainDialog), configuration["ConnectionName"])
        {
            _logger = logger;

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

            //AddDialog(new TextPrompt(nameof(TextPrompt)));
            AddDialog(new ChoicePrompt(nameof(ChoicePrompt)));

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
                    await stepContext.Context.SendActivityAsync(MessageFactory.Text("You are now logged in."), cancellationToken);
                    return await stepContext.NextAsync(null, cancellationToken);
                }
            }

            await stepContext.Context.SendActivityAsync(MessageFactory.Text("Login was not successful please try again."), cancellationToken);
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
                    else if (command.StartsWith("show me"))
                    {
                        await stepContext.Context.SendActivityAsync(
                            MessageFactory.Text("I don't know how to do this yet!"),
                            cancellationToken);
                    }
                    else if (command.StartsWith("show calendar"))
                    {
                        await stepContext.Context.SendActivityAsync(
                            MessageFactory.Text("I don't know how to do this yet!"),
                            cancellationToken);
                    }
                    else if (command.StartsWith("add event"))
                    {
                        await stepContext.Context.SendActivityAsync(
                            MessageFactory.Text("I don't know how to do this yet!"),
                            cancellationToken);
                    }
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
    }
}

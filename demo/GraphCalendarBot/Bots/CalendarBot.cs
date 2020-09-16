// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
//
// Generated with EchoBot .NET Template version v4.10.2

// <CalendarBotSnippet>
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.Teams;
using Microsoft.Bot.Schema;
using Microsoft.Extensions.Logging;

namespace CalendarBot.Bots
{
    public class CalendarBot<T> : TeamsActivityHandler where T : Dialog
    {
        protected readonly BotState ConversationState;
        protected readonly Dialog Dialog;
        protected readonly ILogger Logger;
        protected readonly BotState UserState;

        public CalendarBot(
            ConversationState conversationState,
            UserState userState,
            T dialog,
            ILogger<CalendarBot<T>> logger)
        {
            ConversationState = conversationState;
            UserState = userState;
            Dialog = dialog;
            Logger = logger;
        }

        public override async Task OnTurnAsync(
            ITurnContext turnContext,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            Logger.LogInformation("CalendarBot.OnTurnAsync");
            await base.OnTurnAsync(turnContext, cancellationToken);

            // Save any state changes that might have occurred during the turn.
            await ConversationState.SaveChangesAsync(turnContext, false, cancellationToken);
            await UserState.SaveChangesAsync(turnContext, false, cancellationToken);
        }

        protected override async Task OnMessageActivityAsync(
            ITurnContext<IMessageActivity> turnContext,
            CancellationToken cancellationToken)
        {
            Logger.LogInformation("CalendarBot.OnMessageActivityAsync");
            await Dialog.RunAsync(turnContext,
                ConversationState.CreateProperty<DialogState>(nameof(DialogState)),
                cancellationToken);
        }

        protected override async Task OnMembersAddedAsync(
            IList<ChannelAccount> membersAdded,
            ITurnContext<IConversationUpdateActivity> turnContext,
            CancellationToken cancellationToken)
        {
            Logger.LogInformation("CalendarBot.OnMembersAddedAsync");
            var welcomeText =
                "Welcome to Microsoft Graph CalendarBot. Type anything to get started.";

            foreach (var member in membersAdded)
            {
                if (member.Id != turnContext.Activity.Recipient.Id)
                {
                    await turnContext.SendActivityAsync(
                        MessageFactory.Text(welcomeText),
                        cancellationToken);
                }
            }
        }

        protected override async Task OnTokenResponseEventAsync(
            ITurnContext<IEventActivity> turnContext,
            CancellationToken cancellationToken)
        {
            Logger.LogInformation("CalendarBot.OnTokenResponseEventAsync");
            await Dialog.RunAsync(turnContext,
                ConversationState.CreateProperty<DialogState>(nameof(DialogState)),
                cancellationToken);
        }

        protected override async Task OnTeamsSigninVerifyStateAsync(
            ITurnContext<IInvokeActivity> turnContext,
            CancellationToken cancellationToken)
        {
            Logger.LogInformation("CalendarBot.OnTeamsSigninVerifyStateAsync");
            await Dialog.RunAsync(turnContext,
                ConversationState.CreateProperty<DialogState>(nameof(DialogState)),
                cancellationToken);
        }
    }
}
// </CalendarBotSnippet>

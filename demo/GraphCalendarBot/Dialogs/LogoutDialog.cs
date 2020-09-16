// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

// <LogoutDialogSnippet>
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Schema;

namespace CalendarBot.Dialogs
{
    public class LogoutDialog : ComponentDialog
    {
        public LogoutDialog(string id, string connectionName)
            : base(id)
        {
            ConnectionName = connectionName;
        }

        protected string ConnectionName { get; private set; }

        // All dialogs should inherit this class so the user
        // can log out at any time
        protected override async Task<DialogTurnResult> OnBeginDialogAsync(
            DialogContext innerDc,
            object options,
            CancellationToken cancellationToken)
        {
            // Check if this is a logout command
            var result = await InterruptAsync(innerDc, cancellationToken);
            if (result != null)
            {
                return result;
            }

            return await base.OnBeginDialogAsync(innerDc, options, cancellationToken);
        }

        protected override async Task<DialogTurnResult> OnContinueDialogAsync(
            DialogContext innerDc,
            CancellationToken cancellationToken)
        {
            // Check if this is a logout command
            var result = await InterruptAsync(innerDc, cancellationToken);
            if (result != null)
            {
                return result;
            }

            return await base.OnContinueDialogAsync(innerDc, cancellationToken);
        }

        private async Task<DialogTurnResult> InterruptAsync(
            DialogContext innerDc,
            CancellationToken cancellationToken)
        {
            // If this is a logout command, cancel any other activities and log out
            if (innerDc.Context.Activity.Type == ActivityTypes.Message)
            {
                var text = innerDc.Context.Activity.Text.ToLowerInvariant();

                if (text.StartsWith("log out") || text.StartsWith("logout"))
                {
                    // The bot adapter encapsulates the authentication processes.
                    var botAdapter = (BotFrameworkAdapter)innerDc.Context.Adapter;
                    await botAdapter.SignOutUserAsync(
                        innerDc.Context, ConnectionName, null, cancellationToken);
                    await innerDc.Context.SendActivityAsync(
                        MessageFactory.Text("You have been signed out."), cancellationToken);
                    return await innerDc.CancelAllDialogsAsync();
                }
            }
            return null;
        }
    }
}
// </LogoutDialogSnippet>

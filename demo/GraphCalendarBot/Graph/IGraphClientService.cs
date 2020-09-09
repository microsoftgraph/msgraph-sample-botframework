// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

// <IGraphClientServiceSnippet>
using Microsoft.Graph;

namespace CalendarBot.Graph
{
    public interface IGraphClientService
    {
        GraphServiceClient GetAuthenticatedGraphClient(string accessToken);
    }
}
// </IGraphClientServiceSnippet>

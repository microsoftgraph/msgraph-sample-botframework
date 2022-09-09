// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using Microsoft.Graph;

namespace CalendarBot.Services
{
    public interface IGraphClientService
    {
        GraphServiceClient GetAuthenticatedGraphClient(string accessToken);
    }
}

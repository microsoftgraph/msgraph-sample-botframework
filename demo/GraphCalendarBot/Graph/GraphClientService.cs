// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

// <GraphClientServiceSnippet>
using Microsoft.Graph;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace CalendarBot.Graph
{
    public class GraphClientService : IGraphClientService
    {
        public GraphServiceClient GetAuthenticatedGraphClient(string accessToken)
        {
            return new GraphServiceClient(new DelegateAuthenticationProvider(
                async (request) => {
                    // Add the access token to the Authorization header
                    // on the outgoing request
                    request.Headers.Authorization =
                        new AuthenticationHeaderValue("Bearer", accessToken);
                    await Task.CompletedTask;
                }
            ));
        }
    }
}
// </GraphClientServiceSnippet>

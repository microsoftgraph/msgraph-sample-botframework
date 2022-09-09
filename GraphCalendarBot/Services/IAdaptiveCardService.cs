// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using AdaptiveCards;
using Microsoft.Graph;
using System.IO;

namespace CalendarBot.Services
{
    public interface IAdaptiveCardService
    {
        AdaptiveCard GetUserCard(User user, Stream photo);
        AdaptiveCard GetEventCard(Event calendarEvent, string dateTimeFormat);
    }
}

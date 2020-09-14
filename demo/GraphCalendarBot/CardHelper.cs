// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using AdaptiveCards;
using Microsoft.Graph;
using System;
using System.IO;

namespace CalendarBot
{
    public class CardHelper
    {
        public static AdaptiveCard GetUserCard(User user, Stream photo)
        {
          // Create an Adaptive Card to display the user
            // See https://adaptivecards.io/designer/ for possibilities
            var userCard = new AdaptiveCard("1.2");

            var columns = new AdaptiveColumnSet();
            userCard.Body.Add(columns);

            var userPhotoColumn = new AdaptiveColumn { Width = AdaptiveColumnWidth.Auto };
            columns.Columns.Add(userPhotoColumn);

            userPhotoColumn.Items.Add(new AdaptiveImage {
                Style = AdaptiveImageStyle.Person,
                Size = AdaptiveImageSize.Small,
                Url = GetDataUriFromPhoto(photo)
            });

            var userInfoColumn = new AdaptiveColumn {Width = AdaptiveColumnWidth.Stretch };
            columns.Columns.Add(userInfoColumn);

            userInfoColumn.Items.Add(new AdaptiveTextBlock {
                Weight = AdaptiveTextWeight.Bolder,
                Wrap = true,
                Text = user.DisplayName
            });

            userInfoColumn.Items.Add(new AdaptiveTextBlock {
                Spacing = AdaptiveSpacing.None,
                IsSubtle = true,
                Wrap = true,
                Text = user.Mail ?? user.UserPrincipalName
            });

            return userCard;
        }

        // <GetEventCardSnippet>
        public static AdaptiveCard GetEventCard(Event calendarEvent, string dateTimeFormat)
        {
            // Build an Adaptive Card for the event
            var eventCard = new AdaptiveCard("1.2");

            // Add subject as card title
            eventCard.Body.Add(new AdaptiveTextBlock
            {
                Size = AdaptiveTextSize.Medium,
                Weight = AdaptiveTextWeight.Bolder,
                Text = calendarEvent.Subject
            });

            // Add organizer
            eventCard.Body.Add(new AdaptiveTextBlock
            {
                Size = AdaptiveTextSize.Default,
                Weight = AdaptiveTextWeight.Lighter,
                Spacing = AdaptiveSpacing.None,
                Text = calendarEvent.Organizer.EmailAddress.Name
            });

            // Add details
            var details = new AdaptiveFactSet();

            details.Facts.Add(new AdaptiveFact
            {
                Title = "Start",
                Value = DateTime.Parse(calendarEvent.Start.DateTime).ToString(dateTimeFormat)
            });

            details.Facts.Add(new AdaptiveFact
            {
                Title = "End",
                Value = DateTime.Parse(calendarEvent.End.DateTime).ToString(dateTimeFormat)
            });

            if (calendarEvent.Location != null &&
                !string.IsNullOrEmpty(calendarEvent.Location.DisplayName))
            {
                details.Facts.Add(new AdaptiveFact
                {
                    Title = "Location",
                    Value = calendarEvent.Location.DisplayName
                });
            }

            eventCard.Body.Add(details);

            return eventCard;
        }
        // </GetEventCardSnippet>

        private static Uri GetDataUriFromPhoto(Stream photo)
        {
            // Copy to a MemoryStream to get access to bytes
            var photoStream = new MemoryStream();
            photo.CopyTo(photoStream);

            var photoBytes = photoStream.ToArray();

            return new Uri($"data:image/png;base64,{Convert.ToBase64String(photoBytes)}");
        }
    }
}
---
ms.localizationpriority: medium
---

<!-- markdownlint-disable MD002 MD041 -->

This tutorial teaches you how to build a Bot Framework bot that uses the Microsoft Graph API to retrieve calendar information for a user.

> [!TIP]
> If you prefer to just download the completed tutorial, you can download or clone the [GitHub repository](https://github.com/microsoftgraph/msgraph-training-botframework). See the README file in the **demo** folder for instructions on configuring the app with an app ID and secret.

## Prerequisites

Before you start this tutorial, you should have the following installed on your development machine.

- [.NET Core SDK](https://dotnet.microsoft.com/download)
- [Bot Framework Emulator](https://github.com/microsoft/BotFramework-Emulator/blob/master/README.md)
- [ngrok](https://ngrok.com/)

You should also have either a personal Microsoft account with a mailbox on Outlook.com, or a Microsoft work or school account. If you don't have a Microsoft account, there are a couple of options to get a free account:

- You can [sign up for a new personal Microsoft account](https://signup.live.com/signup?wa=wsignin1.0&rpsnv=12&ct=1454618383&rver=6.4.6456.0&wp=MBI_SSL_SHARED&wreply=https://mail.live.com/default.aspx&id=64855&cbcxt=mai&bk=1454618383&uiflavor=web&uaid=b213a65b4fdc484382b6622b3ecaa547&mkt=E-US&lc=1033&lic=1).
- You can [sign up for the Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program) to get a free Microsoft 365 subscription.
- An Azure subscription. If you do not have one, create a [free account](https://azure.microsoft.com/free/?WT.mc_id=A261C142F) before you begin.

> [!NOTE]
> This tutorial was written with the following versions. The steps in this guide may work with other versions, but that has not been tested.
>
> - .NET Core SDK version 5.0.302
> - Bot Framework Emulator 4.1.3
> - ngrok 2.3.40

## Feedback

Please provide any feedback on this tutorial in the [GitHub repository](https://github.com/microsoftgraph/msgraph-training-botframework).

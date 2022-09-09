---
ms.localizationpriority: medium
---

<!-- markdownlint-disable MD002 MD041 -->

In this exercise, you will create a new Bot Channels registration and an Azure AD web application registration using the Azure Portal.

## Create a Bot Channels registration

1. Open a browser and navigate to the [Azure Portal](https://portal.azure.com). Login using the account associated with your Azure subscription.

1. Select the upper-left menu, then select **Create a resource**.

    ![A screenshot of the Azure Portal menu](images/create-resource.png)

1. On the **New** page, search for `Azure Bot` and select **Azure Bot**.

1. On the **Azure Bot** page, select **Create**.

1. Fill in the required fields. The **Bot handle** field must be unique. Be sure to review the different pricing tiers and select what makes sense for your scenario. If this is just a learning exercise, you may want to select the free option.

1. For **Type of App**, select **User-Assigned Managed Identity**.

1. For **Creation type**, select **Multi Tenant**.

1. Select **Review + create**. Once validation completes, select **Create**.

1. Once deployment has finished, select **Go to resource**.

1. Under **Settings**, select **Configuration**. Select the **Manage** link next to **Microsoft App ID**.

1. Select **New client secret**. Add a description and choose an expiration, then select **Add**.

1. Copy the client secret value before you leave this page. You will need it in the following steps.

    > [!IMPORTANT]
    > This client secret is never shown again, so make sure you copy it now. You will need to enter this value in multiple places so keep it safe.

1. Select **Overview** in the left-hand menu. Copy the value of the **Application (client) ID** and save it, you will need it in the following steps.

## Create a web app registration

1. Return to the home page of the Azure portal, then select **Azure Active Directory**.

1. Select **App registrations**.

1. Select **New registration**. On the **Register an application** page, set the values as follows.

    - Set **Name** to `Graph Calendar Bot Auth`.
    - Set **Supported account types** to **Accounts in any organizational directory and personal Microsoft accounts**.
    - Under **Redirect URI**, set the first drop-down to `Web` and set the value to `https://token.botframework.com/.auth/web/redirect`.

    > **NOTE:** `https://token.botframework.com/.auth/web/redirect` is the default Bot Framework OAuth redirect URL for the public Azure cloud with no data residency requirements. Depending on your environment, you may need to use a different redirect URL. See [OAuth URL support in Azure Bot Service](https://docs.microsoft.com/azure/bot-service/ref-oauth-redirect-urls?view=azure-bot-service-4.0) for more information.

1. Select **Register**. On the **Graph Calendar Bot Auth** page, copy the value of the **Application (client) ID** and save it, you will need it in the following steps.

1. Select **Certificates & secrets** under **Manage**. Select the **New client secret** button. Enter a value in **Description** and select one of the options for **Expires** and select **Add**.

1. Copy the client secret value before you leave this page. You will need it in the following steps.

1. Select **API permissions**, then select **Add a permission**.

1. Select **Microsoft Graph**, then select **Delegated permissions**.

1. Select the following permissions, then select **Add permissions**.

    - **openid**
    - **profile**
    - **Calendars.ReadWrite**
    - **MailboxSettings.Read**

    ![A screenshot of configured permissions](images/configured-permissions.png)

### About permissions

Consider what each of those permission scopes allows the bot to do, and what the bot will use them for.

- **openid** and **profile**: allows the bot to sign users in and get basic information from Azure AD in the identity token.
- **Calendars.ReadWrite**: allows the bot to read the user's calendar and to add new events to the user's calendar.
- **MailboxSettings.Read**: allows the bot to read the user's mailbox settings. The bot will use this to get the user's selected time zone.
- **User.Read**: allows the bot to get the user's profile from Microsoft Graph. The bot will use this to get the user's name.

## Add OAuth connection to the bot

1. Navigate to your bot's Azure Bot page in the Azure Portal. Select **Configuration** under **Settings**.

1. Select **Add OAuth Connection Settings**.

1. Fill in the form as follows, then select **Save**.

    - **Name**: `GraphBotAuth`
    - **Provider**: **Azure Active Directory v2**
    - **Client id**: The application ID of your **Graph Calendar Bot Auth** registration.
    - **Client secret**: The client secret of your **Graph Calendar Bot Auth** registration.
    - **Token Exchange URL**: Leave blank
    - **Tenant ID**: `common`
    - **Scopes**: `https://graph.microsoft.com/.default`

1. Select the **GraphBotAuth** entry under **OAuth Connection Settings**.

1. Select **Test Connection**. This opens a new browser window or tab to start the OAuth flow.

1. If necessary, sign in. Review the list of requested permissions, then select **Accept**.

1. You should see a **Test Connection to 'GraphBotAuth' Succeeded** message.

> [!TIP]
> You can select the **Copy Token** button on this page, and paste the token into [https://jwt.ms](https://jwt.ms) to see the claims inside the token. This is useful when troubleshooting authentication errors.

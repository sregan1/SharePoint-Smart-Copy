# Permissions

Which Microsoft permissions SharePoint Smart Copy needs, and for which feature. You create the app registration in your own tenant, so **you decide what to grant**.

**Smart Copy's core job (copying files, libraries, lists and sites) needs only three permissions, and one of them is for Migration API mode only. Everything else is optional**: add a permission only if you want the feature that needs it.

This page is the master copy. The same text is in the [README](README.md#which-permissions-does-each-feature-need) and the [User Guide](USER-GUIDE.md#which-permissions-does-each-feature-need).

For help working out which permissions your administrator needs to grant, see the [permissions guide](https://sharepointsmartsolutions.com/smart-copy/permissions) on the Smart Copy website.

Add permissions under **Entra ID → App registrations → your registration → API permissions → Add a permission**, then click **Grant admin consent for your organization**.

## Which permissions does each feature need?

**Required: copying files, libraries, lists and sites.** Smart Copy needs only these permissions for its core job: Files & Folders, Libraries & Lists, Entire Site, Site Pages, Compare and verifying a copy. The third one is for **Migration API mode only**: leave it out if you use only Enhanced REST mode. **Everything below this table is optional**: add a permission only if you want the feature that needs it.

| API | Type | Permission |
|---|---|---|
| Microsoft Graph | Delegated | `Sites.ReadWrite.All` |
| Microsoft Graph | Delegated | `Files.ReadWrite.All` |
| SharePoint | Delegated | `AllSites.FullControl` (for Migration API mode; can be excluded if you use only Enhanced REST mode) |

Without the first two nothing connects to SharePoint. Without the third, Migration API mode is refused ("Permission Denied") and Enhanced REST mode still works. Click **Grant admin consent** once they are added.

**Optional: add only what you want.** Each item says what you want to do, which permissions to add *on top of the required ones*, and what happens if you leave them out. Permissions are **Delegated** (the app acts as the signed-in user) unless marked **Application**.

**If you want to copy managed-metadata (term set) columns to a different tenant**

- Add on the **target** registration: Microsoft Graph, Delegated: `TermStore.ReadWrite.All`
- Add on the **source** registration (when two registrations are used): Microsoft Graph, Delegated: `TermStore.Read.All`
- The signed-in target account also needs a term store role (for example *Term Store Administrator*).
- Without it: the managed-metadata column is not created, and the run report says why.

**If you want to use the Forms tool**

- Add: Microsoft Forms, Delegated: the Microsoft Forms permission (under *APIs my organization uses*)
- Without it: the Forms tile is grayed out, and its tooltip names what to add.
- The Forms tool uses Microsoft's unofficial Forms API.

**If you want the Flows tool to rewrite exported flow packages**

- Add nothing: the two required permissions are enough. Adding the Microsoft Forms permission above also lets it look up form titles.

**If you want Flows to copy flows live between environments**

- Add: Power Automate Service, Delegated: `Flows.Read.All`, `Flows.Manage.All`
- Add: PowerApps Service, Delegated: `User`
- Optional, to create the flows inside a solution: Dataverse, Delegated: `user_impersonation`. The signed-in account also needs a security role that can create flows (for example *Environment Maker*).
- Without it: the Live option is disabled and says why. Packages still work.

**If you want to browse your own Teams and chats**

- Add: Microsoft Graph, Delegated: `Team.ReadBasic.All`, `Channel.ReadBasic.All`, `TeamMember.Read.All`, `Chat.Read`
- Without it: the Teams tile shows a warning, and the tool cannot list your teams and chats.

**If you want to copy teams (without messages), in the same tenant or another one**

- Add the browsing permissions above, plus Microsoft Graph, Delegated: `Team.Create`, `Channel.Create`, `ChannelMember.ReadWrite.All`, `TeamMember.ReadWrite.All`, `TeamSettings.ReadWrite.All`
- Between two tenants, the **target** registration also needs Microsoft Graph, Delegated: `User.ReadBasic.All`
- Without it: the wizard stops on its first step and names the missing permissions.

**If you want to copy Teams messages and replies**

- Add the team-copy permissions above, plus Microsoft Graph, Delegated: `ChannelMessage.Read.All`, `ChatMember.Read`
- Add on the **target** registration: Microsoft Graph, **Application**: `Teamwork.Migrate.All`, and a client secret for that registration, saved in Settings
- Without it: the *Include messages* option is disabled (no client secret), or Microsoft refuses the import (no permission).

**If you want to copy chats**

- Add the message permissions above, plus Microsoft Graph, Delegated: `Chat.Create`
- Add on the **target** registration: Microsoft Graph, **Application**: `ChatMember.ReadWrite.All`
- Without it: the new chat cannot be created, or its messages are imported but its people cannot see them yet (the results say so).

**If you want the Teams administrator view: every team in the tenant, and the chats of people you choose**

- Add on the **source** registration: Microsoft Graph, **Application**: `Group.Read.All`, `Team.ReadBasic.All`, `Channel.ReadBasic.All`, `TeamMember.Read.All`, `ChannelMember.Read.All`, `Files.Read.All`, `User.Read.All`
- Also save a client secret for the **source** registration in Settings. The tool then reads with that registration's own sign-in instead of your account.
- To also read other people's chats, or the messages of teams you are not in, add on the source registration: Microsoft Graph, **Application**: `ChannelMessage.Read.All`, `Chat.Read.All`, `ChatMember.Read.All`. Microsoft also requires its approval of your tenant for the Teams *protected APIs* (a request form); admin consent alone is not enough.
- Without it: the Teams tool lists only your own teams and chats, and a message names what is missing.
- Files of a team you are not a member of are copied with your own sign-in, so they copy only where your account can open the team's site.

**Two registrations (two tenants).** Each side uses its own app registration and its own consent. The source registration needs only what is read from the source; the target registration needs what is written there. A registration used for both sides needs both sets. Grant admin consent in **each** tenant.

**Notes**

- **Admin consent.** Application permissions always need admin consent, and so do most of the Teams permissions. The simplest way is to click *Grant admin consent* once after adding everything.
- **Client secret.** Only the Application features above need one: on the target registration to import Teams messages and chats, and on the source registration for the Teams administrator view. Create it under *Certificates & secrets* on that registration and paste its **Value** (not its ID) into **Settings → Azure app registrations**. It is stored encrypted for your Windows account and is never sent anywhere except to Microsoft's sign-in service.
- **Grayed-out tiles.** Forms is disabled until its permission is on the registration. Teams is never disabled: it shows a warning tooltip you can copy and paste to your administrator.
- **No permissions needed** to browse saved History or open Settings (local only).

## Telling your administrator what is missing

- **Forms** and **Teams** tiles: hover the tile for the missing permissions, or right-click it and choose **Copy this message for your administrator** to paste into an email.
- In the **Flows** tool the Live option explains its missing permissions the same way.
- The Teams wizard's first step names any missing copy permission, and the *Include messages* option explains when the client secret is missing.

## Related

- [README: setting up the app registration](README.md#2-add-api-permissions)
- [User Guide: Azure app registration](USER-GUIDE.md#4-setup--azure-app-registration)
- Consent is per tenant. Repeat the setup in each tenant you copy between.

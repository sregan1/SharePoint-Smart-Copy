# Permissions

Which Microsoft permissions SharePoint Smart Copy needs, for which feature, and what you see when one is missing. You create the app registration in your own tenant, so **you decide which of these to grant**: add only the groups for the features you use.

This page is the master copy. The same chart is in the [README](README.md#which-permissions-does-each-feature-need) and the [User Guide](USER-GUIDE.md#which-permissions-does-each-feature-need).

## The chart

Permissions are grouped under a short code so the feature chart stays readable. Add only the groups for the features you use. Everything is **delegated** (the app acts as the signed-in user) except those marked **Application**, which are used only for Teams messages and the Teams administrator view.

**Permission groups**

| Code | API | Type | Permissions | If it is missing |
|---|---|---|---|---|
| **BASE** | Microsoft Graph | Delegated | `Sites.ReadWrite.All`, `Files.ReadWrite.All` | Nothing connects to SharePoint: sign-in or the first request fails |
| **MIG** | SharePoint | Delegated | `AllSites.FullControl` | Migration API mode is refused ("Permission Denied"); Enhanced REST mode still works |
| **TERM** | Microsoft Graph | Delegated | `TermStore.ReadWrite.All` (target), and `TermStore.Read.All` on the source registration when two registrations are used | The managed-metadata column is not created; the run report says why |
| **FORMS** | Microsoft Forms | Delegated | the Microsoft Forms permission (under *APIs my organization uses*) | The Forms tile is grayed out and its tooltip names what to add; in Flows only form-title lookups fail |
| **FLOW** | Power Automate Service, PowerApps Service | Delegated | `Flows.Read.All`, `Flows.Manage.All` (Power Automate Service); `User` (PowerApps Service) | The Live option in Flows is disabled, with the reason; packages still work |
| **DV** | Dataverse | Delegated | `user_impersonation` | Live-copied flows are created outside a solution |
| **TREAD** | Microsoft Graph | Delegated | `Team.ReadBasic.All`, `Channel.ReadBasic.All`, `TeamMember.Read.All`, `Chat.Read` | The Teams tile shows a warning, and the tool cannot list your teams and chats |
| **TWRITE** | Microsoft Graph | Delegated | `Team.Create`, `Channel.Create`, `ChannelMember.ReadWrite.All`, `TeamMember.ReadWrite.All`, `TeamSettings.ReadWrite.All` | The wizard stops on the first step and names the missing permissions |
| **TUSER** | Microsoft Graph | Delegated | `User.ReadBasic.All` | Different-tenant copies cannot look people up in the target and report a permission error |
| **TMSG** | Microsoft Graph | Delegated | `ChannelMessage.Read.All`, `ChatMember.Read` | Reading messages or chat members from the source fails with a permission error |
| **TCHAT** | Microsoft Graph | Delegated | `Chat.Create` | The new chat cannot be created |
| **TAPP** | Microsoft Graph | **Application** | `Teamwork.Migrate.All`, plus the target registration's **client secret** saved in Settings | The *Include messages* option is disabled without the secret; without the permission Microsoft refuses the import |
| **TCHATAPP** | Microsoft Graph | **Application** | `ChatMember.ReadWrite.All` | A copied chat's messages are imported but its people cannot see them yet; the results say so |
| **TADMIN** | Microsoft Graph | **Application** (source registration) | `Group.Read.All`, `Team.ReadBasic.All`, `Channel.ReadBasic.All`, `TeamMember.Read.All`, `ChannelMember.Read.All`, `Files.Read.All`, `User.Read.All`, plus the **source** registration's client secret saved in Settings | Without the secret the Teams tool shows only your own teams and chats; without the permissions the list says which to add |
| **TADMINMSG** | Microsoft Graph | **Application** (source registration) | `ChannelMessage.Read.All`, `Chat.Read.All`, `ChatMember.Read.All`, **and** Microsoft's approval of the tenant for the Teams *protected APIs* (a request form; admin consent alone is not enough) | Reading other people's chats, or the messages of a team you are not in, is refused with a message that names the approval |

**What each feature needs**

| Feature | Required | Optional |
|---|---|---|
| Files & Folders, Libraries & Lists, Entire Site — **Enhanced REST** mode | BASE | TERM (managed metadata to a different tenant) |
| Files & Folders, Libraries & Lists, Entire Site — **Migration API** mode | BASE, MIG | TERM (managed metadata to a different tenant) |
| Site Pages (always Enhanced REST) | BASE | — |
| **Compare**, and verifying a copy | BASE | — |
| **Forms** tool | BASE, FORMS | — |
| **Flows** tool — packages | BASE | FORMS (form-title lookups) |
| **Flows** tool — live copy between environments | BASE, FLOW | FORMS, DV |
| **Teams** — browse teams and chats | BASE, TREAD | — |
| **Teams** — copy a team in the same tenant (settings, members, channels, files) | BASE, TREAD, TWRITE | — |
| **Teams** — copy a team to a different tenant | BASE, TREAD, TWRITE, TUSER (target registration) | — |
| **Teams** — include messages and replies | all of the above for the copy, plus TMSG and TAPP | — |
| **Teams** — copy a chat | BASE, TREAD, TWRITE, TMSG, TCHAT, TAPP, TCHATAPP (and TUSER between tenants) | — |
| **Teams** — administrator view: every team in the tenant, copied without its messages | BASE, TADMIN, TWRITE (and TUSER between tenants) | — |
| **Teams** — administrator view: the chats of chosen people, or messages of teams you are not in | all of the above, plus TADMINMSG, TMSG, TCHAT, TAPP, TCHATAPP | — |
| Browsing saved History, Settings | none (local only; verifying a run signs in like Compare) | — |

**Two tenants:** each side uses its own app registration and its own consent. The source registration needs only what is read from the source (BASE, and for Teams TREAD, and TMSG when copying messages or chats); the target registration needs what is written there (BASE, MIG for Migration API mode, and for Teams TWRITE, TUSER, TCHAT, TAPP and TCHATAPP). A registration used for both sides needs the union. Grant admin consent in **each** tenant.

**Notes**

- **Admin consent.** Application permissions always need admin consent, and so do most of the Teams permissions: simplest is to click *Grant admin consent* once after adding everything.
- **Client secret.** Only TAPP needs one. Create it under *Certificates & secrets* on the target registration, then paste it into **Settings → Azure app registrations**. It is stored encrypted for your Windows account and is never sent anywhere except to Microsoft's sign-in service.
- **Forms** uses Microsoft's unofficial Forms API. In a Dataverse environment the signed-in account needs a security role that can create flows (for example *Environment Maker*). Managed-metadata copy also needs a term store role on the target account (for example *Term Store Administrator*).
- **Grayed-out tiles.** Forms is disabled until its permission is on the registration. Teams is never disabled: it shows a warning tooltip you can copy and paste to your administrator.

## Copy-and-paste lists for common setups

Add these under **Entra ID → App registrations → your registration → API permissions → Add a permission**, then click **Grant admin consent for your organization**.

**1. Copy SharePoint files, libraries and sites (Enhanced REST)**

- Microsoft Graph, Delegated: `Sites.ReadWrite.All`, `Files.ReadWrite.All`

**2. As above, with Migration API mode (version history, faster large copies)**

- everything in 1
- SharePoint, Delegated: `AllSites.FullControl`

**3. Forms**

- everything in 1
- Microsoft Forms, Delegated: the Microsoft Forms permission

**4. Flows, packages only**

- everything in 1 (Forms permission optional)

**5. Flows, live copy**

- everything in 1
- Power Automate Service, Delegated: `Flows.Read.All`, `Flows.Manage.All`
- PowerApps Service, Delegated: `User`
- optional, to create flows inside a solution: Dataverse, Delegated: `user_impersonation`

**6. Teams, browse only**

- everything in 1
- Microsoft Graph, Delegated: `Team.ReadBasic.All`, `Channel.ReadBasic.All`, `TeamMember.Read.All`, `Chat.Read`

**7. Teams, copy teams (no messages)**

- everything in 6
- Microsoft Graph, Delegated: `Team.Create`, `Channel.Create`, `ChannelMember.ReadWrite.All`, `TeamMember.ReadWrite.All`, `TeamSettings.ReadWrite.All`
- between two tenants, on the target registration also: `User.ReadBasic.All`

**8. Teams, copy teams with messages**

- everything in 7
- Microsoft Graph, Delegated: `ChannelMessage.Read.All`, `ChatMember.Read`
- Microsoft Graph, **Application**: `Teamwork.Migrate.All`
- a client secret on the **target** registration (Certificates & secrets), saved in Settings

**9. Teams, copy chats**

- everything in 8
- Microsoft Graph, Delegated: `Chat.Create`
- Microsoft Graph, **Application**: `ChatMember.ReadWrite.All`

**10. Managed metadata to a different tenant**

- Microsoft Graph, Delegated: `TermStore.ReadWrite.All` on the target registration, and `TermStore.Read.All` on the source registration when two registrations are used
- the signed-in target account also needs a term store role (for example *Term Store Administrator*)

**Administrator view (Teams).** Without a client secret the Teams tool lists only the teams and chats of the account you signed in with. With a client secret saved for the source registration the tool lists **every team in the tenant** and lets you add **the chats of people you choose**; tick *Only show my own teams and chats* on the Connect step for the personal view. It reads with the *source* registration's own application sign-in (its client secret, saved in Settings), not with your account, so the groups TADMIN (teams, channels, members, files, people) and, for messages and chats, TADMINMSG are needed on the source registration, with admin consent in the source tenant. Microsoft requires a separate approval (its *protected APIs* request form) before an application may read messages and chats; until it is granted the tool shows a message that says so. Files of a team you are not a member of are copied with your own sign-in, so they are copied only where your account can open the team's site.

## Telling your administrator what is missing

- **Forms** and **Teams** tiles: hover the tile for the missing permissions, or right-click it and choose **Copy this message for your administrator** to paste into an email.
- In the **Flows** tool the Live option explains its missing permissions the same way.
- The Teams wizard's first step names any missing copy permission, and the *Include messages* option explains when the client secret is missing.

## Related

- [README: setting up the app registration](README.md#2-add-api-permissions)
- [User Guide: Azure app registration](USER-GUIDE.md#4-setup--azure-app-registration)
- Consent is per tenant. Repeat the setup in each tenant you copy between.

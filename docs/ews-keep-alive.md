# Keeping Exchange Web Services (EWS) on past October 1, 2026

A stopgap for Microsoft 365 admins. It buys time, at most until **April 1, 2027**. It is not a permanent fix.

## Why this matters

Starting October 1, 2026, Microsoft is switching off Exchange Web Services (EWS) by default in Microsoft 365 tenants, and removing it permanently on April 1, 2027. Anything that still relies on EWS stops connecting. That includes legacy ("classic") Outlook for Mac, Apple Mail Exchange accounts on macOS 26 and earlier, and many third-party tools.

A tenant admin can keep EWS on until April 1, 2027.

## Before you start

- The setting is organization-wide and can **only** be changed in PowerShell. The Exchange admin center doesn't expose it. There is a per-mailbox EWS toggle (mailbox → *Manage email apps*), but it can't override the organization setting.
- It may **not** rescue legacy Outlook for Mac. Microsoft is ending that app's support on its own schedule and may block it regardless of this setting.
- You need an Exchange or Global admin account, and Windows PowerShell (or PowerShell 7 on a Mac).

## Steps (Windows)

1. Start menu → type `PowerShell` → right-click → **Run as administrator**.

2. Install the Exchange Online module (one time only):
   ```powershell
   Install-Module ExchangeOnlineManagement -Scope CurrentUser
   ```
   - "NuGet provider is required": answer **Y**.
   - "Untrusted repository … PSGallery": answer **A** (Yes to All). PSGallery is Microsoft's official module gallery; it's labeled untrusted only because it hasn't been marked trusted on that PC.
   - If it says running scripts is disabled, run this, then retry the install:
     ```powershell
     Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
     ```

3. Sign in with your admin account:
   ```powershell
   Connect-ExchangeOnline -UserPrincipalName admin@yourdomain.com
   ```
   You may get a popup saying "Device management could not be enabled." Click **OK**. That's Windows trying to enroll the PC in your organization's device management; it isn't needed here. If the sign-in fails, use plain browser sign-in instead:
   ```powershell
   Connect-ExchangeOnline -UserPrincipalName admin@yourdomain.com -DisableWAM
   ```
   If you manage several tenants, run `Disconnect-ExchangeOnline` before switching, and check which account you're on with `Get-ConnectionInformation`.

4. Record the current setting before changing it:
   ```powershell
   Get-OrganizationConfig | Format-List Name, EwsEnabled, EwsAllowedAppIDs
   ```
   - **Blank** means never set. Microsoft hasn't switched your tenant off yet.
   - **False** means it's already been switched off. Step 5 turns it back on, but expect a short outage until it takes effect.

5. Turn EWS on, confirm, and sign out:
   ```powershell
   Set-OrganizationConfig -EwsEnabled:$true
   Get-OrganizationConfig | Format-List EwsEnabled, EwsAllowedAppIDs
   Disconnect-ExchangeOnline
   ```
   You should see `EwsEnabled : True`. Leave `EwsAllowedAppIDs` empty. Filling it in restricts EWS to only the listed apps, and you'd need the exact app IDs for everything that uses it.

## To undo it

```powershell
Set-OrganizationConfig -EwsEnabled $null
```

That puts it back to the default (blank), after which Microsoft's switch-off applies.

## How to tell whether it worked

Over the following days, compare Apple Mail (with the account set up as **Exchange**, not IMAP) against legacy Outlook for Mac:

| Apple Mail (Exchange) | Legacy Outlook | What it means |
|---|---|---|
| Works | Works | The setting held. You have until April 1, 2027. |
| Works | Stops | Microsoft blocked Outlook for Mac separately. Move those users to new Outlook. |
| Stops | Stops | The setting didn't hold. Re-check step 4. |

## Longer term

Use the time to move everything off EWS before April 1, 2027: new Outlook for Mac, and Microsoft Graph–based versions of any third-party tools, such as the [AccountEdge Graph Mailer](../README.md) for AccountEdge emailing.

Reference: [Exchange Online EWS, Your Time is Almost Up (Microsoft)](https://techcommunity.microsoft.com/blog/exchange/exchange-online-ews-your-time-is-almost-up/4492361)

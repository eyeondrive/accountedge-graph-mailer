# AccountEdge Graph Mailer

Keep **AccountEdge for Mac** emailing invoices, statements and other forms after Microsoft shuts off the old way Outlook for Mac talks to Microsoft 365.

It sends AccountEdge's email **directly through Microsoft 365** (the Microsoft Graph API). No Outlook, no Apple Mail, nothing to keep scriptable. Your users keep clicking **Send To → Email** in AccountEdge exactly as they do today.

> **Not affiliated with Acclivity/AccountEdge or Microsoft.** Use at your own risk, and test it on your own setup before relying on it for billing. See [Disclaimer](#disclaimer).

---

## Why this exists

AccountEdge for Mac emails forms by handing them to a helper that drives your mail app, usually **legacy ("classic") Outlook for Mac**, with AppleScript. That chain is breaking:

- **Microsoft is retiring Exchange Web Services (EWS).** It's switched off by default from **October 1, 2026** and removed for good on **April 1, 2027**. Legacy Outlook for Mac depends on EWS to reach Microsoft 365 mailboxes, and Microsoft has separately announced the end of legacy Outlook for Mac.
- **New Outlook for Mac can't take legacy Outlook's place.** Microsoft has cancelled AppleScript support for new Outlook, so AccountEdge can't drive it.
- **Apple Mail isn't a dependable fallback on older Macs.** Its Exchange support on macOS 26 and earlier still relies partly on EWS. Apple's move of Mail to Microsoft Graph is tied to macOS 27, which Intel Macs can't run.

This tool removes the mail app from the chain entirely.

> **Need time first?** A Microsoft 365 admin can keep EWS switched on until April 1, 2027. See [docs/ews-keep-alive.md](docs/ews-keep-alive.md). That's a stopgap, not a fix.

---

## How it works

When you email a form, AccountEdge writes a small text file, `~/Library/Application Support/OfficeLink/OfficeLink Data`, then launches a helper app (`OfficeLink Processor.app`) to act on it. The same helper also handles a few non-email jobs, such as an "are you there?" check.

The installer swaps the helper's script for a small dispatcher:

- **Email jobs** (`AppID=MYOB.EMAIL`) go to `send.sh`, which sends through Microsoft Graph.
- **Everything else** is handed, unchanged, to an untouched copy of the **original** helper.

`send.sh` handles every email the same way, so every send exercises the whole path:

1. Create the message as a draft in the sender's mailbox.
2. Attach each file. Files under 3 MB go in one request; larger ones are uploaded in pieces, up to your tenant's message size limit (35 MB by default).
3. Send it. Microsoft files a copy in the sender's **Sent Items**.

Success is silent. Any failure shows an on-screen dialog, and the email is not sent. If a listed attachment is missing, the email is refused rather than sent without it.

<details>
<summary>The <code>OfficeLink Data</code> format (worked out by observation)</summary>

An INI-style text file with Windows (CRLF) line endings. **Recipient and attachment keys repeat**, one line per entry:

| Key | Meaning |
|---|---|
| `AppID` | Job type. `MYOB.EMAIL` = send email; `MYOB.EXISTS` = "are you there?" check |
| `Subject` | Subject line |
| `Content` | Message body (plain text) |
| `Address` | To (repeatable) |
| `CCAddress` | Cc (repeatable) |
| `BCAddress` | Bcc (repeatable) |
| `Attachment` | Full path to a file (repeatable; the form PDF plus any extra attachments) |
| `EmailAction` | `Send` |

```
[Settings]
AppID=MYOB.EMAIL
[Fill]
Subject=Invoice 1234
Content=Please find your invoice attached.
Address=customer@example.com
CCAddress=accounts@example.com
Attachment=/Users/you/Library/Application Support/AccountEdge Pro/Bottles/.../Inv00001234.pdf
EmailAction=Send
```
</details>

---

## Requirements

- AccountEdge for Mac that has emailed at least once on the machine (so the OfficeLink helper exists). Tested with the CrossOver-based AccountEdge Pro. See [Tested on](#tested-on).
- macOS with its built-in tools only (`bash`, `curl`, `plutil`). Nothing to install.
- A **Microsoft 365 / Exchange Online** mailbox to send from.
- **Global Administrator** on that Microsoft 365 tenant, for one-time setup.
- PowerShell for one setup step: Windows PowerShell, or PowerShell 7 on a Mac.

---

## Setup

Do steps 1–2 **once per Microsoft 365 tenant** (once per organization). Do steps 3–5 **on each Mac**, signed in as the Mac user who runs AccountEdge.

### 1. Register the app in Microsoft Entra

In the **Entra admin center** ([entra.microsoft.com](https://entra.microsoft.com)). This is not the Intune or Microsoft 365 admin center.

1. **Entra ID → App registrations → New registration.** Name it `AccountEdge Graph Mailer`. Supported account types: **single tenant**. Leave the redirect URI blank. **Register**.
2. From the **Overview** page, copy the **Application (client) ID** and the **Directory (tenant) ID**.
3. **Certificates & secrets → New client secret.** Pick an expiry (e.g. 24 months) and **put a reminder in your calendar to rotate it**. When it expires, sending stops. Copy the **Value** right away; it's shown only once. Don't paste it into email or notes.
4. **API permissions → Add a permission → Microsoft Graph → Application permissions** (not Delegated). Add **`Mail.Send`** and **`Mail.ReadWrite`**.
5. **Grant admin consent.** Click **Grant admin consent for _your org_** on the API permissions page. If you don't see that button, go to **Enterprise applications → AccountEdge Graph Mailer → Permissions → Grant admin consent**, and accept. The popup lists them as "Send mail as any user" and "Read and write mail in all mailboxes".

> **Why both permissions?** `Mail.Send` sends. `Mail.ReadWrite` is required to build the message as a draft and attach files over 3 MB. Without it, emails would be capped at about 3 MB of attachments, and AccountEdge gives users no warning about attachment size. Step 2 confines both permissions to the mailbox(es) you choose.

### 2. Restrict the app to your sending mailbox(es)

**Don't skip this.** Without it, these permissions apply to **every mailbox in your organization**.

1. In the **Exchange admin center**, create a **mail-enabled security group** (e.g. `AccountEdge Senders`) and add the mailbox(es) AccountEdge will send from. Using a group means you can add or remove senders later without touching the app.
2. In PowerShell:

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser   # one time
Connect-ExchangeOnline -UserPrincipalName admin@yourdomain.com

New-ApplicationAccessPolicy -AppId <CLIENT_ID> -PolicyScopeGroupId accountedge-senders@yourdomain.com `
  -AccessRight RestrictAccess -Description "AccountEdge Graph Mailer - allowed senders only"

Test-ApplicationAccessPolicy -Identity billing@yourdomain.com -AppId <CLIENT_ID>   # expect: Granted
Test-ApplicationAccessPolicy -Identity someone.else@yourdomain.com -AppId <CLIENT_ID>   # expect: Denied
Disconnect-ExchangeOnline
```

A single-person organization can put one mailbox address in `-PolicyScopeGroupId` instead of a group. The policy can take 30+ minutes to take effect.

### 3. Store the secret in the Mac's Keychain

Keychain keeps the secret encrypted, so it never sits in a plain text file. Use the Mac user account that runs AccountEdge; **not** a root shell (a prompt ending in `#`), or it goes into root's keychain where the tool can't find it.

**Option A, Terminal.** Copy the secret Value, then:

```bash
pbpaste | wc -c                                   # should print about 40, the secret's length; it doesn't show the secret
security add-generic-password -s "aegm-yourorg" -a "$USER" -T /usr/bin/security -w "$(pbpaste)"
pbcopy < /dev/null                                # clear the clipboard
```

**Option B, Keychain Access app.** Select the **login** keychain → **File → New Password Item**. Item name `aegm-yourorg`, account name = your Mac username, password = the secret. Then run this once in Terminal and click **Always Allow**, so sends don't prompt later:

```bash
security find-generic-password -s "aegm-yourorg" -w > /dev/null
```

> If macOS offers a **key icon / AutoFill** suggestion while you paste, ignore it. It offers your *saved website passwords*, not the client secret. Option A avoids the prompt entirely.

### 4. Create the config

```bash
cp config.example.yaml my-config.yaml
```

Fill in `tenant_id`, `client_id`, `keychain_item`, and `sender`. Leave `client_secret` blank when using Keychain.

### 5. Install

```bash
./install.sh --config my-config.yaml
```

The installer:
- backs up the original helper **once** to `~/Library/Application Support/accountedge_graph_mailer/backup/` (never overwritten);
- places a working copy of the original helper beside it for non-email jobs;
- deploys the scripts and your config (locked to your user, `chmod 600`);
- swaps the helper's script for the dispatcher, then re-signs it.

`./install.sh --dry-run` shows the plan without changing anything.

---

## Test it

```bash
# 1. Sign-in and permissions check (sends nothing)
~/Library/Application\ Support/accountedge_graph_mailer/check.sh
#    expect: permissions : ["Mail.ReadWrite","Mail.Send"]  ...  ✓ ready to send

# 2. Dry run against the last thing AccountEdge wrote (sends nothing)
~/Library/Application\ Support/accountedge_graph_mailer/dispatch.sh \
  --config ~/Library/Application\ Support/accountedge_graph_mailer/config.yaml --dry-run
```

3. In AccountEdge, email an invoice **to yourself**, including a large extra attachment (over 3 MB). Check that it arrives, with all attachments, and appears in the sender's **Sent Items**. Also send one copy to an **outside** address (e.g. Gmail) and confirm it lands in the inbox, not spam.

Every action is logged to `~/Library/Logs/accountedge_graph_mailer.log`:

```
2026-10-01 11:35:32 dispatch AppID=MYOB.EXISTS -> original applet
2026-10-01 11:38:31 dispatch AppID=MYOB.EMAIL -> send.sh
2026-10-01 11:38:40 OK sent subj='Invoice 1234' to=customer@example.com cc= bcc= att=1 (0.1 MB)
```

---

## Uninstall

```bash
./uninstall.sh
```

Restores the original helper from the backup, so AccountEdge goes back to emailing through Outlook. It leaves `~/Library/Application Support/accountedge_graph_mailer/` (your config) in place. Delete that yourself, and the Keychain item, if you're removing it for good.

---

## Troubleshooting

| Symptom | Likely cause / fix |
|---|---|
| `check.sh`: `permissions : none` or a permission missing | Permission added as **Delegated** instead of **Application**, or admin consent not granted. See setup step 1. |
| "Grant admin consent" button missing | Use **Enterprise applications → the app → Permissions** instead. Granting needs **Global Administrator**. |
| A permission shows in App registrations but not Enterprise applications | It hasn't been consented yet. Grant admin consent. |
| Sign-in failed / `invalid_client` | Wrong client ID, or the secret is wrong or **expired**. Create a new secret and update the Keychain item. |
| Send fails with HTTP 403 | The sender isn't allowed by the Application Access Policy (or the policy is still propagating), or consent is missing. |
| "Microsoft would not send… Total attachments: N MB" | Attachments exceed the tenant's message size limit (35 MB default). |
| Keychain item "not found" | It was created from a root shell or under a different Mac user. Recreate it as the user who runs AccountEdge. |
| "Device management could not be enabled" during `Connect-ExchangeOnline` | Harmless: Windows tried to enroll the PC. Click OK; if sign-in fails, add `-DisableWAM`. |
| PowerShell "Untrusted repository PSGallery" | Microsoft's official module gallery. Answer **A** (Yes to All). |
| Email works, but another AccountEdge OfficeLink feature stopped | Check the log for `-> original applet` lines and an `rc=` other than 0, and open an issue. `./uninstall.sh` restores everything. |
| An AccountEdge update reverted the helper | Re-run `./install.sh --config …`. The original backup is preserved. |

---

## Security notes

- The secret lives in the user's **login Keychain**. The config holds only the Keychain item's name.
- `Mail.ReadWrite` + `Mail.Send` are powerful **application** permissions. The **Application Access Policy** (setup step 2) confines them to the mailboxes you list. Treat the client secret like a password to those mailboxes.
- Use **one app registration per Microsoft 365 tenant**. If you support several organizations, don't reuse one multi-tenant app and secret across them: a secret copied off one client's Mac shouldn't open another organization's mailboxes.
- Secrets expire. Calendar the rotation.

---

## Limitations

- **Tested on:** see below. Native (non-CrossOver) AccountEdge builds use the same OfficeLink helper but are **untested**.
- Only the "are you there?" check (`MYOB.EXISTS`) has been observed among non-email jobs. Other OfficeLink jobs (e.g. Send to Excel/Word) pass to the original helper but are untested.
- The message body is sent as **plain text**, as AccountEdge provides it.
- Failures are reported on screen only. `alert_to` is reserved for future email alerts.
- An AccountEdge update may restore the original helper. Re-run `install.sh` if emailing reverts to Outlook.

### Tested on

- macOS 15.7 (Sequoia), Apple Silicon
- AccountEdge Pro (CrossOver-based build), October 2026
- Microsoft 365 / Exchange Online

Reports from other setups are welcome. Please open an issue.

---

## Repository layout

```
install.sh / uninstall.sh          install and restore
config.example.yaml                annotated config template
script/dispatch.sh                 routes OfficeLink jobs: email -> send.sh, others -> original helper
script/send.sh                     sends via Microsoft Graph (draft -> attachments -> send)
script/check.sh                    sign-in and permissions check (sends nothing)
script/officelink_applet.applescript   replacement helper script (calls dispatch.sh)
docs/ews-keep-alive.md             keeping EWS on until April 2027 (admin stopgap)
```

---

## Disclaimer

Provided as-is under the MIT License, without warranty. This project isn't affiliated with, endorsed by, or supported by Acclivity Group LLC (AccountEdge) or Microsoft. "AccountEdge", "Microsoft 365", "Outlook" and related names are trademarks of their owners. No AccountEdge software is included; the installer works only with each user's own copy. Test on your own setup before relying on it for billing.

## License

[MIT](LICENSE)

---
layout: post
title: "SharePoint Online file locks: diagnose first, recover checkouts with PowerShell"
date: 2026-10-23 12:00:00 -0400
categories: [sharepoint, powershell]
permalink: /sharepoint/repair-spo-file-lock/
---

A “locked for editing” message is a symptom, not a diagnosis. Repair-SPOFileLock.ps1 v2.1 helps distinguish explicit checkouts from reported editing locks, collect evidence, and recover eligible checkouts with safeguards.

<!--more-->

<p><a href="https://github.com/pvutla001/pvutla001.github.io/blob/main/scripts/Repair-SPOFileLock.ps1">View the complete PowerShell source on GitHub</a> · <a href="/scripts/Repair-SPOFileLock.ps1" download="Repair-SPOFileLock.ps1">Download Repair-SPOFileLock.ps1</a></p>

**Author: Prem V · Script version: 2.1**

## Start with the type of lock

When a document will not open or save in SharePoint Online, the first task is to establish what SharePoint reports. An explicit checkout and an Office editing-session lock require different recovery paths.

This script reads the file's checkout type, checkout owner, and reported lock owner. It can also scan a document library recursively, producing a report of the files visible to the signed-in account.

It is not a universal “force unlock” tool. The `Unlock` action is an alias for `CheckIn`: it checks in an explicit checkout when the prerequisites are met. It does not release an arbitrary Office editing lock.

## What the script does

| Action | Behavior |
|---|---|
| `Diagnose` | Default. Reads the target's reported state without changing it in SharePoint. |
| `WaitForRelease` | Polls a single file for a clear observation within a specified interval. |
| `CheckIn` | Attempts check-in of an explicit checkout after confirmation and backup safeguards. |
| `Unlock` | Maps to `CheckIn`; it is not an editing-lock force release. |
| `DiscardCheckout` | Explicitly discards checkout changes, subject to additional acknowledgement and safeguards. |

Library scans accept only `Diagnose`. They do not perform bulk repair or download file content.

## Before running it

You need access to the target SharePoint web, a compatible installation of `PnP.PowerShell`, and permissions appropriate to the requested operation. Being able to read a file does not mean you can check in another user's checkout.

The script declares PowerShell 5.1 as its minimum, but your installed PnP module has its own runtime requirements. That declaration does not make current PnP releases compatible with Windows PowerShell 5.1. Use a PowerShell host supported by your chosen module version.

The default `Interactive` path uses browser-based authentication. Provide your own Entra application client ID, or configure a default client ID as described in the [PnP client-ID guidance](https://pnp.github.io/powershell/articles/defaultclientid.html). See [PnP authentication documentation](https://pnp.github.io/powershell/articles/authentication.html) for application configuration.

The optional `LegacyWebLogin` branch imports PnP.PowerShell **1.12.0** and requires **Windows PowerShell 5.1 in STA mode**. The script labels this an unsupported legacy authentication method; tenant policy may reject it. Treat it as a compatibility option, not the preferred setup.

The code uses SharePoint CSOM and REST, with no explicit Microsoft Graph calls or device-code authentication.

**Run it as a script, not by dot-sourcing it.** It ends with `exit` and sets a process exit code. Use a separate PowerShell process when you need to isolate that behavior.

## 1. Diagnose a single document

Download the script, save it as `Repair-SPOFileLock.ps1`, and start with a diagnosis:

```powershell
./Repair-SPOFileLock.ps1 `
    -FileUrl "https://contoso.sharepoint.com/sites/QA/Documents/Test.docx" `
    -ClientId "<your-app-client-id>"
```

All examples use placeholder values. Replace them with your own approved test resource.

For a direct file URL, omit sharing and viewer parameters. A literal `#` in the filename must be encoded as `%23`.

Alternatively, use `-SiteUrl` with `-FilePath`. The latter expects a **decoded server-relative path**, such as `/sites/QA/Documents/Test.docx`; do not pass a sharing link or a URL-encoded path.

For subsites or nonstandard URL structures, supply `-SiteUrl` explicitly. The script verifies that the connection points to the requested SharePoint web.

## 2. Scan a library

```powershell
./Repair-SPOFileLock.ps1 `
    -LibraryUrl "https://contoso.sharepoint.com/sites/QA/Documents" `
    -ClientId "<your-app-client-id>"
```

The scan walks visible files recursively across file types. Library view filters are ignored. Review the CSV, then investigate affected files individually.

A completed enumeration is limited to what the account can see. Failed file reads remain unknown; they are not evidence of an unlocked file.

## 3. Close editing sessions and observe release

Ask the user to save their work and close relevant Office and browser sessions, including sessions on other devices. Then observe the file:

```powershell
./Repair-SPOFileLock.ps1 `
    -FileUrl "https://contoso.sharepoint.com/sites/QA/Documents/Test.docx" `
    -ClientId "<your-app-client-id>" `
    -Action WaitForRelease `
    -ClientSessionsClosed `
    -WaitSeconds 600 `
    -PollSeconds 15
```

The defaults are 600 seconds of waiting and 15 seconds between polls. A clear final observation means no checkout or editing lock was reported at that instant. It does not establish the age of a lock, a lease expiry, or proof that a persistent lock is orphaned.

## 4. Preview an explicit checkout recovery

Only use `-ClientSessionsClosed` after confirming that editing sessions are closed and saved changes have been reconciled.

```powershell
./Repair-SPOFileLock.ps1 `
    -FileUrl "https://contoso.sharepoint.com/sites/QA/Documents/Test.docx" `
    -ClientId "<your-app-client-id>" `
    -Action CheckIn `
    -ClientSessionsClosed `
    -WhatIf
```

`-WhatIf` prevents remote changes, but the script still authenticates, reads the file, and writes local audit reports. It does not exercise the backup or prove that a subsequent check-in will succeed.

After reviewing the diagnosis and preview, run the same command without `-WhatIf` to request the actual recovery. Keep the confirmation prompt.

Before mutation, the script:

1. Downloads server-readable file bytes to a local backup and records its SHA-256 hash.
2. Confirms the backup length matches the observed file length.
3. Reads the state again and compares identity, ETag, modification time, size, checkout owner, and lock owner.
4. Stops if the state changed during preparation.
5. Attempts check-in, then reads the resulting state.

It chooses minor check-in when the library supports minor versions; otherwise it chooses major check-in. Check-in is not the same as approval or publication.

The backup does **not** preserve unsaved client edits, the complete version history, or all metadata. The pre-mutation comparison reduces races; it does not make the operation atomic.

## Discard is a separate decision

`DiscardCheckout` requires both `-ClientSessionsClosed` and `-AcknowledgeDiscardChanges`. It is blocked when no major version exists, protecting initial uploads and unpublished drafts from that discard path.

Discarding checkout changes can lose work even when a server-content backup exists. Use it only after deciding that discarding those changes is appropriate.

## Citrix and other remote desktops

The script accepts incident context through `-ClientEnvironment`, `-ClientSessionState`, `-SessionHost`, and `-SessionId`. These are **operator-supplied values**, not discovered session information.

A disconnected or unreachable session needs attention from the desktop or VDI administrator so edits can be preserved. The script does not reconnect, log off, terminate, or save work from remote sessions.

Do not combine `-ClientSessionsClosed` with a non-`Unknown` session state: the script rejects that conflicting input before proceeding.

## Reports and exit codes

By default, each run creates a unique UTC-stamped folder under `SPO-Lock-Reports` beside the script. Use `-OutputDirectory` to select another writable location.

Outputs include a log, JSON report, CSV report, and text summary. An approved mutation also creates a before-change content backup. Reports may contain user identities, file paths, incident details, and errors; store them with appropriate access controls.

| Exit code | Interpretation |
|---|---|
| `0` | A clear observation, completed clear scan, or successful eligible outcome; inspect the status for the exact meaning. |
| `1` | Failure, incomplete file reads, or a mutation/verification error requiring review. |
| `2` | A checkout, lock signal, inconsistent state, or remaining signal requiring investigation. |
| `3` | A blocked operation, cancelled confirmation, or WhatIf with no remote changes. |

After a mutation timeout or verification failure, review the latest server state before retrying: the server may already have committed the operation.

## What a clear report cannot prove

No lock signal does not guarantee that a document will open, save, or publish. Permissions, application validation, format compatibility, records or retention restrictions, and client problems need separate investigation.

For a persistent editing lock, preserve work, verify session closure, wait and diagnose again, then provide the collected evidence and actual error to Microsoft support if needed.

## Validation status

This article documents the supplied v2.1 source. The publication review checked its behavior and examples; **the script has not been executed against a SharePoint tenant as part of preparing this post**. Validate it in a test library before operational use.

The source header mentions a companion `README-v2.txt` that was not included with this upload. This post provides usage guidance for the supplied script; it does not claim to reproduce that missing document.

<p><a href="https://github.com/pvutla001/pvutla001.github.io/blob/main/scripts/Repair-SPOFileLock.ps1">View source</a> · <a href="/scripts/Repair-SPOFileLock.ps1" download="Repair-SPOFileLock.ps1">Download the script</a></p>

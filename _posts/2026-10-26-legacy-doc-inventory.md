---
layout: post
title: "Inventory legacy Word documents in SharePoint Online with PowerShell"
date: 2026-10-26 12:00:00 -0400
categories: [sharepoint, powershell]
permalink: /sharepoint/legacy-doc-inventory/
excerpt: "Find legacy .doc files, export metadata and folder counts, and optionally inspect VBA indicators using a four-file PowerShell solution."
---

Before a document migration or format-conversion project, you need to know where legacy Word files are, how many remain, and which may need special handling. **Export-SPOLegacyDocInventory v3.1** inventories the current `.doc` files visible to your account in one SharePoint Online library and produces Excel reports plus a full-value JSONL journal.

This is a file and metadata inventory. It does **not** report document views, opening frequency, or other usage analytics. It does not convert documents or repair locks.

## When this helps

- Scope a legacy Word cleanup or migration before deciding which files to convert.
- Give content owners a list of `.doc` files, locations, modification dates, sizes, and selected metadata.
- Count legacy files by folder to prioritize review.
- Optionally download files and inspect local VBA project indicators before choosing a conversion approach.

Metadata and count modes do not download document contents. Content checks download and retain mirrored local copies, then inspect them without starting Word or executing macros.

## Download the four-file solution

Save **all four files in the same folder**, retaining the filenames below. There is no Python, Microsoft Word, or ImportExcel requirement.

| File | Purpose | Source | Download |
| --- | --- | --- | --- |
| `Export-SPOLegacyDocInventory.ps1` | Entry point, authentication, orchestration and reports | [View](https://github.com/pvutla001/pvutla001.github.io/blob/main/scripts/legacy-doc-inventory/Export-SPOLegacyDocInventory.ps1) | <a href="/scripts/legacy-doc-inventory/Export-SPOLegacyDocInventory.ps1" download="Export-SPOLegacyDocInventory.ps1">Download</a> |
| `SPOLegacyDocInventory.Engine.ps1` | Paging, metadata workers, downloads and scanner process | [View](https://github.com/pvutla001/pvutla001.github.io/blob/main/scripts/legacy-doc-inventory/SPOLegacyDocInventory.Engine.ps1) | <a href="/scripts/legacy-doc-inventory/SPOLegacyDocInventory.Engine.ps1" download="SPOLegacyDocInventory.Engine.ps1">Download</a> |
| `SPOLegacyDocInventory.Runtime.cs` | C# runtime compiled locally for retry handling, file I/O and streaming XLSX output | [View](https://github.com/pvutla001/pvutla001.github.io/blob/main/scripts/legacy-doc-inventory/SPOLegacyDocInventory.Runtime.cs) | <a href="/scripts/legacy-doc-inventory/SPOLegacyDocInventory.Runtime.cs" download="SPOLegacyDocInventory.Runtime.cs">Download</a> |
| `Inspect-LegacyWord.ps1` | PowerShell parser for optional local Word/VBA inspection | [View](https://github.com/pvutla001/pvutla001.github.io/blob/main/scripts/legacy-doc-inventory/Inspect-LegacyWord.ps1) | <a href="/scripts/legacy-doc-inventory/Inspect-LegacyWord.ps1" download="Inspect-LegacyWord.ps1">Download</a> |

If your browser opens a file as text, use **Save link as** and preserve its `.ps1` or `.cs` extension. Avoid saving an HTML page as a script.

The public launcher has no configured tenant URL. Supply your own `-LibraryUrl` explicitly. All SharePoint URLs and application IDs in this article are placeholders.

## Prerequisites and authentication

Use **Windows PowerShell 5.1 in STA mode**, a writable script/output folder, network access to SharePoint Online, and an account permitted to read the target library. Content checks also need permission to download the files. Files hidden by permissions are outside the report.

This implementation pins **PnP.PowerShell 1.12.0**, a legacy version. It is not designed to run with the current PnP module or PowerShell 7. If that exact version is missing, the launcher can save it and its dependencies under `.dependencies`; this needs access to PowerShell Gallery and may install the NuGet provider for the current user. Use `-NoInstall` to prevent automatic dependency installation.

The default `LegacyWebLogin` path uses `-UseWebLogin`. Tenant policy may block this legacy sign-in method. `SystemBrowser` uses interactive authentication and requires your own approved Entra application client ID and appropriate configuration/consent. Neither option bypasses MFA or Conditional Access. See the [PnP authentication guidance](https://pnp.github.io/powershell/articles/authentication.html); current guidance does not make the pinned legacy module a supported modern runtime.

Start a fresh Windows PowerShell process if another PnP module version is already loaded. Follow your organization’s script execution policy.

## 1. Create a metadata inventory

From the folder containing the four files, run:

```powershell
powershell.exe -NoProfile -STA -File .\Export-SPOLegacyDocInventory.ps1 `
  -LibraryUrl "https://contoso.sharepoint.com/sites/Example/Documents" `
  -ThrottleLimit 4 `
  -BatchSize 500
```

This is the default mode: it discovers `.doc` files recursively at all folder depths and exports their details and selected default/custom metadata. `.docx`, `.docm`, historical versions and deleted files are not included. Library view filters in a supplied URL are ignored.

Use the actual library root URL. For a subweb, supply `-SiteUrl` explicitly. `-Folder` accepts a **decoded library-relative path**, not a URL or server-relative path:

```powershell
powershell.exe -NoProfile -STA -File .\Export-SPOLegacyDocInventory.ps1 `
  -SiteUrl "https://contoso.sharepoint.com/sites/Example" `
  -LibraryUrl "https://contoso.sharepoint.com/sites/Example/Documents" `
  -Folder "Policies/Finance" `
  -OutputDirectory "C:\Reports\LegacyWord"
```

## 2. Count files by folder

`-CountOnly` creates folder counts instead of detailed document metadata. It defaults to a maximum folder depth of **2**, measured from the library root. To include every depth:

```powershell
powershell.exe -NoProfile -STA -File .\Export-SPOLegacyDocInventory.ps1 `
  -LibraryUrl "https://contoso.sharepoint.com/sites/Example/Documents" `
  -CountOnly `
  -MaxDepth -1
```

Use `-MaxDepth 2` for a bounded report. `DirectDocCount` counts files directly in that folder; `DocsWithinDepthLimit` includes descendant files within the chosen depth limit. Do not add parent and child aggregate counts together. A scoped folder deeper than `MaxDepth` is rejected. `-CountOnly` cannot be combined with `-DownloadFiles`.

## 3. Inspect VBA and macro indicators

To inventory, download and statically inspect the matching files:

```powershell
powershell.exe -NoProfile -STA -File .\Export-SPOLegacyDocInventory.ps1 `
  -LibraryUrl "https://contoso.sharepoint.com/sites/Example/Documents" `
  -CheckBoth `
  -MaxScanMB 100 `
  -ScanTimeoutSeconds 120
```

`-CheckVBA`, `-CheckMacro` and `-CheckBoth` are mutually exclusive modes. All three use the same native parser; VBA and macro indicators overlap. `-CheckBoth` populates both sets of report columns; it is not two independent scanners.

The parser recognizes Word 97+ OLE binaries and VBA project/source indicators in the outer compound-file hierarchy. WordBasic, embedded packed files, external templates and compiled-only/stomped VBA are not exhaustively assessed. Encrypted, unsupported, malformed, oversized or timed-out files require review. **NotDetected is not a malware-free verdict, and Unknown is not clean.**

The default content limit is 100 MB per file and the scanner timeout is 120 seconds. Content modes retain local downloads under a unique run folder. `-DownloadFiles` can also be used on its own to download without VBA inspection; the content size limit still applies.

## 4. Use system-browser authentication

Where your approved application and tenant configuration support the pinned module:

```powershell
powershell.exe -NoProfile -STA -File .\Export-SPOLegacyDocInventory.ps1 `
  -LibraryUrl "https://contoso.sharepoint.com/sites/Example/Documents" `
  -AuthenticationMode SystemBrowser `
  -ClientId "<your-approved-application-client-id>" `
  -NoInstall
```

Replace the client-ID placeholder with your own application GUID. This example assumes PnP.PowerShell 1.12.0 and its dependencies are already available. If authentication fails, review application configuration and tenant policy rather than treating the failure as an empty inventory.

## Read the output

By default each run creates a unique timestamped folder under `DocInventoryReports` beside the launcher. `-OutputDirectory` changes the parent location. Downloads, when requested, go under `DocInventoryDownloads` unless `-DownloadDirectory` is supplied.

| Output | What to review |
| --- | --- |
| `Inventory_<run>_<part>.xlsx` | `Documents`, `Metadata` and `Fields` worksheets in detailed modes. |
| `FolderCounts_<run>_<part>.xlsx` | Folder-level `.doc` counts in count mode. |
| `RunSummary_<run>.xlsx` and `Summary.json` | Run status, coverage, counts, failures and report parts. |
| `Inventory.jsonl` | Streamed records with full metadata values; Excel cell values may be truncated. |
| `Progress.json` | Latest progress snapshot, worker activity, retries and errors. |
| `Run_<run>.log` | Timestamped progress and diagnostic messages. |

Inspect `Status`, `EnumerationComplete`, `Errors`, `FailedBatches`, and per-document metadata/download/scan status before relying on totals. Partial reports and the journal can survive a failed run; the presence of a workbook alone does not establish completeness.

The script reads SharePoint content and writes local outputs; it does not convert, update, check in or delete source documents. Generated reports can contain account names, file URLs, IDs and metadata from **your** environment, and downloads contain document content. Keep those outputs private; the downloadable source files contain generalized examples.

## How it handles larger libraries

Discovery pages through library items ordered by ID, without a server-side extension filter. The script selects `.doc` records locally and reads metadata in bounded field batches. One discovery request and up to `ThrottleLimit` batch workers use separate cloned CSOM contexts. Shared backoff handles retryable throttling responses.

JSONL and worksheet XML are streamed to disk; Excel report parts rotate rather than holding a complete library workbook in memory. Defaults are 10,000 documents or 200,000 metadata rows per inventory workbook. For a cautious initial run, scope a test folder and reduce `-ThrottleLimit` to 1 or 2.

A starting ID high-watermark excludes newly added IDs above that point. Concurrent moves, edits or deletions can still affect the scan. This is **not a transactional snapshot or a resumable synchronization engine**. Content-change detection marks affected rows for review, but does not eliminate every race.

## Validate before operational use

This publication was reviewed against the four supplied v3.1 source files and sanitized before upload. It has **not been executed against a SharePoint tenant as part of preparing this article**. Test metadata, count and optional content modes in a small library with known examples, and compare the outputs with those files before a wider run.

[Browse all four source files](https://github.com/pvutla001/pvutla001.github.io/tree/main/scripts/legacy-doc-inventory) · [Return to articles](/#articles)

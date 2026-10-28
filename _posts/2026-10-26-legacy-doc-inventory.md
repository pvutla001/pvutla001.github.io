---
layout: post
title: "Inventory legacy Word documents in SharePoint Online with PowerShell v3.10.1"
date: 2026-10-26 12:00:00 -0400
last_modified_at: 2026-10-28 08:34:48 -0400
categories: [sharepoint, powershell]
tags: [legacy-documents, vba, macros, metadata, folder-counts, word97plusole, mhtml, flat-opc]
permalink: /sharepoint/legacy-doc-inventory/
excerpt: "Count legacy .doc files without downloads, export useful metadata, inspect more Word formats for VBA indicators, and review detailed Excel reports with the four-file v3.10.1 solution."
---

*A practical guide to Export-SPOLegacyDocInventory.ps1: folder counts, useful metadata, format identification, VBA indicators and detailed Excel reports.*

**Updated October 28, 2026 · Version 3.10.1.** This update fixes Windows PowerShell 5.1 path-validation failures. The article remains at the same address; download all four matching files together.

A library can contain thousands of files named `.doc` without every file using the same underlying format. Some are Word binary documents, some contain VBA, and others are RTF, XML or web exports saved with a `.doc` extension. A filename inventory tells you where those files are. Content inspection adds evidence about what those files contain and which ones need further review.

`Export-SPOLegacyDocInventory.ps1` brings those tasks into one PowerShell workflow. It can count legacy documents without downloading them, export selected metadata, retain local copies with their folder hierarchy, and inspect supported file structures for VBA indicators. It produces timestamped Excel reports, detailed diagnostics and a final summary of coverage and outcomes.

Version 3.10 expands the formats that the scanner understands and improves its handling of unfamiliar or damaged input. A document that cannot be fully inspected receives an explicit review result. Later documents can continue through the pipeline, and known positive VBA evidence is retained even if another part of the same document fails inspection.

This article uses a fictional Contoso library in the commands. Replace its URL and example folder names with your own. The local installation path shown is `C:\Tools\SPOLegacyDocInventory\V3.10`; change it if you extracted the release elsewhere.

## What this tool helps you investigate

A migration team can identify folders with the highest concentration of legacy documents. A records team can collect titles, ownership-related metadata and modification dates. A support team can distinguish ordinary Word binary files from renamed XML or RTF files. Administrators can separate complete inspections from encrypted, unsupported or incomplete results before deciding what to convert or investigate.

A high `.doc` count identifies a concentration of legacy files; it does not prove those files are broken. Likewise, detecting VBA does not establish that VBA caused a Word Online error. The report supports triage and planning. It does not convert documents, execute macros, repair Word files or release editing locks.

## v3.10.1: fix for Windows PowerShell path errors

Version **3.10.1** fixes repeated download failures such as `Exception calling "CreateDirectory": Illegal characters in path`. The v3.10 engine passed a `\\?\` device path to a .NET Framework string-path API. Hosts using legacy path validation could reject that prefix in `FileIOPermission`, even for a short, valid destination. The failure occurred while preparing the local download directory, before Word format or VBA inspection.

The correction covers directory creation, file existence checks, download streams, partial-file renaming and cleanup, snapshot length/hash reads, mirrored-path checks, scanner input/output and path-identity comparison. Document I/O now uses Windows Unicode file APIs and handle-based .NET streams, preserving long paths without changing the registry or `powershell.exe.config`. Microsoft describes the [Framework compatibility behavior](https://learn.microsoft.com/en-us/dotnet/framework/migration-guide/mitigation-path-normalization), [Unicode directory API](https://learn.microsoft.com/en-us/windows/win32/api/fileapi/nf-fileapi-createdirectoryw) and [handle-based FileStream constructor](https://learn.microsoft.com/en-us/dotnet/api/system.io.filestream.-ctor?view=netframework-4.8).

**Upgrade all four runtime files together, then start a fresh Windows PowerShell 5.1 STA process.** Commands and dependencies remain the same. The summary identifies version `3.10.1`, and the log records the host's path settings without changing them. Windows file failures include the operation, path and Win32 error number. No Python, additional module, separate C# SDK or app registration is introduced by this fix.

Long-path support here applies to retained document files, including deep `Mirror` layouts. Keep the script, module and report roots short because the launcher and reporting code still use standard PowerShell/.NET path APIs; for example, use `C:\Tools\SPOInventory` and `-OutputDirectory 'C:\SPOReports'`. Windows/filesystem component limits and access controls still apply. The former `v3.10.zip` download address remains a compatibility alias for the corrected package.

## Download v3.10.1: four required files

The main script loads its helpers automatically; you normally execute only the entry point.

<p><a class="pv-button pv-primary" style="color:#fff" href="/scripts/legacy-doc-inventory/Export-SPOLegacyDocInventory-v3.10.1.zip" download="Export-SPOLegacyDocInventory-v3.10.1.zip">Download v3.10.1 ZIP</a> <a class="pv-button" href="https://github.com/pvutla001/pvutla001.github.io/tree/main/scripts/legacy-doc-inventory">Browse the source</a></p>

| Required file | Purpose | Source | Download |
| --- | --- | --- | --- |
| `Export-SPOLegacyDocInventory.ps1` | Main script: authentication, progress and reports | [View](https://github.com/pvutla001/pvutla001.github.io/blob/main/scripts/legacy-doc-inventory/Export-SPOLegacyDocInventory.ps1) | <a href="/scripts/legacy-doc-inventory/Export-SPOLegacyDocInventory.ps1" download="Export-SPOLegacyDocInventory.ps1">Download</a> |
| `SPOLegacyDocInventory.Engine.ps1` | Discovery, metadata, workers and downloads | [View](https://github.com/pvutla001/pvutla001.github.io/blob/main/scripts/legacy-doc-inventory/SPOLegacyDocInventory.Engine.ps1) | <a href="/scripts/legacy-doc-inventory/SPOLegacyDocInventory.Engine.ps1" download="SPOLegacyDocInventory.Engine.ps1">Download</a> |
| `SPOLegacyDocInventory.Runtime.cs` | .NET helper for Excel output, contexts and retry handling | [View](https://github.com/pvutla001/pvutla001.github.io/blob/main/scripts/legacy-doc-inventory/SPOLegacyDocInventory.Runtime.cs) | <a href="/scripts/legacy-doc-inventory/SPOLegacyDocInventory.Runtime.cs" download="SPOLegacyDocInventory.Runtime.cs">Download</a> |
| `Inspect-LegacyWord.ps1` | Static format and VBA inspection | [View](https://github.com/pvutla001/pvutla001.github.io/blob/main/scripts/legacy-doc-inventory/Inspect-LegacyWord.ps1) | <a href="/scripts/legacy-doc-inventory/Inspect-LegacyWord.ps1" download="Inspect-LegacyWord.ps1">Download</a> |

The ZIP contains **these four production files and four optional reference files**: `README.md`, `FORMAT-COVERAGE-v3.10.md`, `VALIDATION-v3.10.json` and `SHA256SUMS.txt`. It contains no test fixtures, development builders, historical run logs, downloaded SharePoint documents or PnP module binaries. Tests, fixture builders and historical run results are optional development resources; the four production files do not load them.

If a browser displays a script as text, use **Save link as** and retain its `.ps1` or `.cs` extension. Do not save a GitHub HTML page as the script. The runtime files match the validated v3.10.1 patch; do not mix them with v3.1 or another helper version.

PnP.PowerShell 1.12.0 is an additional module dependency. The script can obtain it if permitted, or you can pre-stage it as described below. New report/download folders are created automatically. Existing `DocInventoryReports`, `DocInventoryDownloads`, `FormatFixtures` and `Validation` folders are not deployment requirements.

## Dependencies and authentication

The production solution uses Windows PowerShell and built-in .NET. Run it in Windows PowerShell 5.1 with STA enabled, preferably in a 64-bit process. The account needs read access to the intended SharePoint library, files and selected metadata, plus write access to the local report and download locations.

The pinned SharePoint dependency is `PnP.PowerShell` **1.12.0**. That exact version preserves this script's legacy authentication workflow. If it is missing, the script uses PowerShellGet to save it into `.dependencies` beside the scripts. It can also install the NuGet package provider for the current user if needed. Use `-NoInstall` to prevent that installation path.

The `.cs` file is source compiled through PowerShell's `Add-Type` and the Windows .NET Framework. No separate C# SDK or Visual Studio installation is needed. ActiveMime inspection also compiles a small input-stream adapter through `Add-Type`. Application controls must therefore permit the script's compilation and child PowerShell processes.

Python, Python packages, desktop Word, desktop Excel and ImportExcel are not required for production inventory or analysis. Excel reports are written directly as `.xlsx` packages. Word is used only by optional fixture builders when recreating Word-authored test samples.

## Run your first inventory

Start in a fresh Windows PowerShell session. This avoids conflicts with another PnP version or an assembly loaded by an older release. From Run, a terminal or an existing shell, start:

```powershell
powershell.exe -NoProfile -STA
```

In that new Windows PowerShell window, run:

```powershell
Set-Location 'C:\Tools\SPOLegacyDocInventory\V3.10'

$library = 'https://contoso.sharepoint.com/sites/Finance/Shared%20Documents'
```

The normal sign-in mode is `LegacyWebLogin`. Complete the sign-in and MFA shown by the script with an account that can read the library. A successful browser session elsewhere does not necessarily authenticate this PowerShell session. Tenant authentication policies can block the retained legacy method; this release does not change those policies.

### Metadata only

Begin with metadata when you need an inventory without file downloads. This is the default mode and includes all folder depths:

```powershell
.\Export-SPOLegacyDocInventory.ps1 -LibraryUrl $library
```

The following explicit form selects the same mode:

```powershell
.\Export-SPOLegacyDocInventory.ps1 -LibraryUrl $library -FetchDocOnly
```

The metadata export retains useful defaults such as Name, Title, Created, Created By, Modified, Modified By, File Type, File Size and Version. It also selects visible custom columns, including custom calculated/read-only fields that pass the script's field-selection rules. Most internal SharePoint fields and hidden custom fields are excluded.

Document identity, URL, folder depth and processing status are still included in the document report because they are needed to locate and interpret results. Selected business metadata is exported separately with field definitions. This is not a dump of every internal SharePoint property.

Metadata-only mode cannot determine whether a document contains VBA. A library's File Type metadata is also different from the actual format identified by reading the file bytes.

### Counts by folder, without downloads

Use count mode to locate folders with many legacy documents. The following command counts `.doc` files at the root and within folders up to two levels below it, without downloading their contents:

```powershell
.\Export-SPOLegacyDocInventory.ps1 `
    -LibraryUrl $library `
    -CountOnly `
    -MaxDepth 2
```

Depth is measured from the library root:

```text
Shared Documents/                     depth 0
Shared Documents/Policies/            depth 1
Shared Documents/Policies/Finance/    depth 2
Shared Documents/Policies/Finance/Old/ depth 3
```

At depth 2, files directly inside `Finance` are included; files inside `Old` are excluded. The summary reports depth exclusions. The default count-mode depth is 2, so specify `-MaxDepth -1` when you want every depth:

```powershell
.\Export-SPOLegacyDocInventory.ps1 `
    -LibraryUrl $library `
    -CountOnly `
    -MaxDepth -1
```

In the folder report, `DirectDocCount` is the number immediately inside a folder. `DocsWithinDepthLimit` includes documents in the permitted descendants. Use direct counts for an additive folder breakdown: summing every descendant-inclusive total would count the same documents more than once.

`-MaxDepth` belongs to count mode. Metadata and content-inspection modes traverse all selected descendants without a fixed six-level limit. The chosen depth limits which files contribute to counts; it does not turn discovery into a traversal that stops querying the library at that depth.

### Download and inspect VBA indicators

Inspect content when you need format and VBA evidence. `-CheckBoth` downloads each selected document, retains a local copy and runs the static parser. This example preserves the hierarchy and enables per-document verbose output:

```powershell
.\Export-SPOLegacyDocInventory.ps1 `
    -LibraryUrl $library `
    -CheckBoth `
    -DownloadLayout Mirror `
    -ThrottleLimit 4 `
    -BatchSize 100 `
    -Verbose
```

The inspection does not start Word, execute VBA or create an editing session in Office. It reads a downloaded snapshot. Metadata and count modes avoid content downloads; content inspection needs the bytes and therefore downloads files.

Choose one reporting mode at a time:

| Command option | Result | Downloads document contents? |
| --- | --- | --- |
| No mode switch, or `-FetchDocOnly` | Metadata inventory | No, unless `-DownloadFiles` is explicitly added |
| `-CountOnly` | Counts by folder | No |
| `-DownloadFiles` with metadata mode | Metadata and retained local copies, without content inspection | Yes |
| `-CheckVBA` | Format inspection and VBA findings | Yes |
| `-CheckMacro` | Format inspection and macro-indicator findings | Yes |
| `-CheckBoth` | Both VBA and macro-indicator report fields | Yes |

`-CheckVBA`, `-CheckMacro` and `-CheckBoth` are alternative parameter sets. Do not combine them. They use the same static parser; they are not separate engines offering different levels of security analysis.

A Word macro is commonly implemented as VBA. Consequently, the `Macro` and `Both` fields describe VBA indicators rather than proving the existence of two independent features. `VBAIndicatorsDetected` does not mean the script executed a callable macro or proved that it would run in Word.

### Target a specific folder

Narrow the work to one folder when investigating a particular area. Supply a decoded path relative to the library root, using forward slashes:

```powershell
.\Export-SPOLegacyDocInventory.ps1 `
    -LibraryUrl $library `
    -Folder 'Policies/Finance/Archived Reviews' `
    -CheckBoth `
    -DownloadLayout Mirror `
    -Verbose
```

This selects that folder and its descendants. Do not pass a full URL or a Windows path to `-Folder`. Use spaces as spaces, rather than converting them to `%20` in this parameter.

Folder selection narrows the documents processed; discovery still enumerates library items and applies the scope. Do not assume a small folder selection eliminates the cost of paging a large library.

For folder counts within a selected branch, depth is still measured from the library root. An unlimited-depth command avoids accidentally excluding a deeply nested target:

```powershell
.\Export-SPOLegacyDocInventory.ps1 `
    -LibraryUrl $library `
    -Folder 'Policies/Finance/Archived Reviews' `
    -CountOnly `
    -MaxDepth -1
```

### Download files without analysis

Download without analysis when you only need retained copies. This keeps metadata collection and content inspection as separate decisions:

```powershell
.\Export-SPOLegacyDocInventory.ps1 `
    -LibraryUrl $library `
    -DownloadFiles `
    -DownloadLayout Mirror
```

`Mirror` retains library-relative folders and maps Windows-invalid names safely. The default `Flat` layout instead uses staging names such as `ItemId_GUID.doc`. Both layouts record the original SharePoint filename, library path, URL and actual `LocalPath`. Use that mapping when comparing reports with local files.

## Repeatable runs and restricted servers

Use splatting for repeatable runs with several options. It keeps commands readable and avoids fragile line-continuation backticks:

```powershell
$inventoryParams = @{
    LibraryUrl      = $library
    CheckBoth       = $true
    DownloadLayout  = 'Mirror'
    OutputDirectory = 'D:\SPOInventory\Reports'
    DownloadDirectory = 'D:\SPOInventory\Downloads'
    ThrottleLimit   = 4
    BatchSize       = 100
    ProgressSeconds = 5
    Verbose         = $true
}

.\Export-SPOLegacyDocInventory.ps1 @inventoryParams
```

The output-directory parameters identify parent directories. The script adds a separate timestamped run directory below each. Ensure the chosen locations have sufficient space for retained documents and reports.

You can also set `$ConfiguredLibraryUrl` inside the main script, then omit `-LibraryUrl`. A supplied parameter takes precedence. For repeatable operations, keeping the URL in a separate calling command or splatting hashtable preserves the distributed script's checksum.

The script accepts a library root URL or a library `Forms/AllItems.aspx` URL. Query-string view filters are ignored. A URL copied from a filtered browser view is therefore not a reliable way to restrict scope; use `-Folder`. For a library in a nested SharePoint subweb, supply the exact web URL:

```powershell
.\Export-SPOLegacyDocInventory.ps1 `
    -SiteUrl 'https://contoso.sharepoint.com/sites/Finance/Operations' `
    -LibraryUrl 'https://contoso.sharepoint.com/sites/Finance/Operations/Shared%20Documents'
```

### Pre-stage dependencies

Pre-stage the pinned module on servers where installation is restricted. On an approved machine with access to PSGallery, save the dependency to a distribution folder:

```powershell
Save-Module -Name PnP.PowerShell `
    -RequiredVersion 1.12.0 `
    -Repository PSGallery `
    -Path 'C:\Staging\SPOInventoryDependencies'
```

Copy the saved module tree and its dependencies, preserving their structure, into `.dependencies` beside the four production files on the target machine. Follow your organization's software distribution process. The directory layout should include `.dependencies\PnP.PowerShell\1.12.0\`.

Then disable installation for the run:

```powershell
.\Export-SPOLegacyDocInventory.ps1 `
    -LibraryUrl $library `
    -CheckBoth `
    -NoInstall `
    -Verbose
```

`-NoInstall` does not make the SharePoint run offline. Network access to the tenant and Microsoft authentication services is still needed. It simply prevents the script from obtaining missing modules/providers. A dependency must be available either in the script's `.dependencies` directory or through the normal module search path.

An alternative authentication mode is present for environments with an approved interactive-login application. It requires an actual non-empty application client ID; the script does not create a registration or grant consent:

```powershell
# Enter your organization's approved interactive application client ID.
$approvedClientId = Read-Host 'Approved application client ID'

.\Export-SPOLegacyDocInventory.ps1 `
    -LibraryUrl $library `
    -AuthenticationMode SystemBrowser `
    -ClientId $approvedClientId `
    -NoInstall
```

This is an optional configured path, not a promise that legacy browser authentication will work under every tenant policy. The recorded live v3.10 validation used the retained legacy login.

## Supported formats and Word 97+ coverage

Version 3.10 identifies content by its structure. Word exposes binary, RTF, HTML, web archive, XML, OpenXML, Flat XML and text save options; filenames alone do not describe all those representations. Microsoft's [Word save-format enumeration](https://learn.microsoft.com/en-us/office/vba/api/word.wdsaveformat) provides the format definitions used when preparing the Word-exported test fixtures.

The library inventory still selects filenames ending in `.doc`, case-insensitively. It does not expand discovery to every `.docx`, `.docm`, `.dot`, `.xml` or `.mht` file. Alternative-format support matters when the selected `.doc` file contains one of those other representations.

| Reported format | What the scanner can establish |
| --- | --- |
| `Word97PlusOLE` | Word 97+ binary compound format; VBA storages and compressed source streams within the supported scope |
| `RTF` | Plain RTF and recognized embedded OLE1/native Word objects |
| `WordOpenXML` | ZIP-based Word documents/templates, including VBA parts and recognized embeddings; tested with macro-enabled and Strict exports |
| `Word2003XML` | Word 2003 XML with recognized binary/ActiveMime VBA and embedded payloads |
| `WordFlatOPC` | OpenXML represented as a single XML file, including recognized VBA and embedding parts |
| `MHTML` | Word web archive with recognized contained Office parts, including ActiveMime payloads |
| `HTML` | HTML structure and references to Office sidecars, scripts or objects; unresolved references require review |
| `PlainText` | Validated simple text/Unicode exports; source-like content or markup remains unresolved |
| `ActiveMime` | Recognized compressed Office storage, decoded within size/time limits and checked for checksum consistency |
| `Word6Or95OLE` / `PreWord97OLE` | Older binary identification; WordBasic inspection is not implemented |
| `EncryptedOLE` or encrypted Word content | Encryption/protection is identified where supported; contents are not decrypted |
| `NotOLE`, unsupported XML/ZIP, PDF or malformed data | Identification and diagnostic evidence where possible; no unsupported negative VBA verdict |

`Word97PlusOLE` is a normal format label. It is not an exception and does not mean the file was created in 1997. Modern Word can also save a legacy binary document.

`FIBVersion` records the base binary revision. `EffectiveFIBVersion` accounts for the newer revision field where present, and `IsTemplate` reports the binary template flag. The base and extension layout is defined by Microsoft's [MS-DOC FIB specification](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/9aeaa2e7-4a45-468e-ab13-3f6193eb9394). These fields describe the file's binary structure, not its creation date or the exact installed Word release that last edited it.

`NotOLE` means the file did not match the recognized compound-file structure or another more specific supported signature. It does not mean macros are absent. Version 3.10 identifies more alternative containers before reaching that result, while retaining uncertainty for unrecognized input.

An HTML file may store its Office information in a separate file such as `editdata.mso`. The scanner records that reference and leaves the result unresolved rather than automatically fetching a URL or opening a sidecar. By contrast, recognized Office data contained inside an MHTML archive can be inspected locally.

## Interpret VBA, macro and review results

Read the findings alongside their inspection scope. `DetectionScope` and `FormatDetails` explain what the parser assessed. The status fields have distinct meanings:

| Field or value | Interpretation |
| --- | --- |
| `Outcome = Succeeded` | The requested operations completed; the document may still contain VBA |
| `Outcome = NeedsReview` | An unsupported/partial inspection, diagnostic, missing metadata or correlation issue needs attention |
| `Outcome = Failed` | An operation such as metadata retrieval or downloading failed |
| `ScanStatus = Scanned` | The requested supported inspection completed for the local snapshot |
| `ScanStatus = Partial` | Some evidence was recovered, but inspection was incomplete |
| `ScanStatus = Unknown` or `Unsupported` | The parser cannot make a complete supported finding |
| `VBA = Detected` | VBA project indicators were found |
| `VBA = NotDetected` | No VBA indicators were found within the stated scope |
| `Macro = VBAIndicatorsDetected` | The static parser found VBA-related evidence; code was not executed |
| `Macro = NoVBAIndicatorsDetected` | No VBA evidence was found within the stated scope |
| `VBAProjectCount` blank | A project count could not be established |

For partial positive results, a known project/module count can be a lower bound. A damaged source stream should not erase the fact that a VBA project was found. Conversely, an unsupported file should not receive a misleading zero count simply because its contents could not be examined.

The parser does not exhaustively assess WordBasic, external templates, packed embedded packages, compiled-only or stomped VBA, or every application-specific interpretation trick. `NotDetected` is not a safety certificate, and `Scanned` does not guarantee that Word Online will open the file.

## Excel reports, diagnostics and retained files

Each run produces a record of both success and uncertainty. By default, outputs are created beside the script:

```text
V3.10/
  DocInventoryReports/
    LegacyDocInventory_<timestamp-and-run-id>/
      RunSummary_<run-id>.txt
      RunSummary_<run-id>.json
      RunSummary_<run-id>.xlsx
      AllDocuments_Metadata_<run-id>_Part0001.xlsx
      SuccessfulDocuments_<run-id>_Part0001.xlsx
      ReviewAndFailures_<run-id>_Part0001.xlsx
      ExecutionLog_<run-id>.log
      FailureDetails_<run-id>.log
      DocumentDiagnostics_<run-id>.jsonl
      DocumentRecords_<run-id>.jsonl
      Progress_<run-id>.json
  DocInventoryDownloads/
    <timestamp-and-run-id>/
      <retained files, when requested>
```

This illustrates a document-processing run. Count mode produces a `FolderCounts` workbook instead of individual document inventory rows. Additional numbered Excel parts are created when report thresholds are reached.

Open `RunSummary` first to check the scope, completion status and coverage. Then use `AllDocuments_Metadata` for the full inventory, `SuccessfulDocuments` for completed operations, and `ReviewAndFailures` for unresolved or failed documents. The review workbook includes a Diagnostics worksheet. An empty category has a header-only workbook, so the existence of a failure workbook alone does not mean failures occurred.

Detailed errors include the file identity, URL, local path, stage, exception type, message and stack. Parser-specific diagnostics can also include a project, stream, module and byte offset. Very long error details may exceed an Excel cell's capacity; the text and JSONL logs retain the fuller evidence.

Errors/review details are recorded without `-Verbose`. Adding `-Verbose` also records successful documents in the structured diagnostic stream and prints their processing results. Use it when investigating individual files; expect larger logs on large libraries.

There is no separate duplicate directory containing only failed documents. If a download succeeded, the review row's `LocalPath` points to its retained copy. A failed download cannot provide a complete local copy.

For a document-processing run, reconcile the counts as:

```text
Succeeded + NeedsReview + Failed = Documents processed
NotProcessed describes matched documents left unprocessed.
```

Enumeration counts include library files and folders, not just selected `.doc` documents. Folder selection and file-extension selection explain why those counts can exceed the processed count. Count-only mode does not claim to have performed individual content inspections.

`CorrelationStatus` is separate from the local finding. The script records a SHA256 hash and compares SharePoint timestamp/ETag evidence around the download and inspection. If a server version changes, the retained local finding remains valid for the recorded bytes, while the correlation warning tells you that the server copy needs another look. `Verified` is version-correlation evidence, not a safety verdict or a transactional guarantee for the entire library.

## Large libraries, progress and performance

Parallelism is bounded so the tool can manage large inventories. Discovery pages through list items in ID order. A starting highest ID bounds the run; a worker pool processes pages with separate SharePoint contexts. Excel rows are streamed to disk and split across workbooks rather than retaining one library-sized workbook in memory. The folder-count map still grows with the number of folders.

| Parameter | Default | What it controls |
| --- | ---: | --- |
| `ThrottleLimit` | 4 | Maximum processing workers; allowed range 1–8 |
| `BatchSize` | 500 | Discovery/processing page size; allowed range 10–2,000 |
| `FieldBatchSize` | 8 | Metadata fields requested per field batch; allowed range 1–8 |
| `ProgressSeconds` | 5 | Progress update interval; allowed range 1–60 seconds |
| `RetryCount` | 6 | Maximum retries for eligible retryable requests |
| `ReportDocumentsPerWorkbook` | 10,000 | Document threshold for report part rotation |
| `ReportMetadataRowsPerWorkbook` | 200,000 | Metadata-row threshold for inventory report rotation |
| `MaxScanMB` | 100 | Per-file content limit in MiB, also applied to requested downloads |
| `ScanTimeoutSeconds` | 120 | Maximum duration of an isolated inventory scanner process |

Four workers is a reasonable starting point for this release's configured workflow. If throttling, CPU pressure or disk contention increases, reduce concurrency and inspect progress/retry evidence. SharePoint can return HTTP 429 or 503 with a `Retry-After` delay; increasing request pressure can prolong throttling. Microsoft's [SharePoint throttling guidance](https://learn.microsoft.com/en-us/sharepoint/dev/general-development/how-to-avoid-getting-throttled-or-blocked-in-sharepoint-online) explains those responses. The script coordinates backoff across its workers when it can recover eligible retry information.

A conservative content run can use two workers and smaller batches:

```powershell
.\Export-SPOLegacyDocInventory.ps1 `
    -LibraryUrl $library `
    -CheckBoth `
    -ThrottleLimit 2 `
    -BatchSize 50 `
    -ProgressSeconds 5 `
    -NoInstall
```

The discovery request is additional to the processing workers. `ThrottleLimit` is not a promise of an exact HTTP request rate, and larger batches do not guarantee faster content analysis. Completed counts are committed when a worker finishes its batch, so counts may remain unchanged while workers are active. Smaller content batches can make completed-result updates arrive sooner.

Content inspection is slower than counting because it reads metadata, transfers bytes, writes and hashes a local copy, starts a separate parser process, examines structures and checks server correlation. Some XML/MHTML payloads also require decoding or decompression. The scanner performs that analysis locally; it does not call a public file-analysis service.

Internal bounds remain in effect even when the outer file limit is increased. Non-OLE container paths have a 32 MiB outer limit; decoded entries have size limits, aggregate embedded content is bounded, and object count/nesting and XML/CFB structures have additional checks. A limit result belongs in review. Increasing `-MaxScanMB` alone cannot make every large or complicated document inspectable.

The architecture targets large libraries, including 200,000+ items, but there is **no measured 200,000-file live benchmark** in this release's validation. Interactive authentication, tenant load, number of custom fields, file sizes and local I/O all affect duration. It is also not a resumable synchronization engine: rerunning starts a new run. Current-version visibility, permission restrictions, files changing during the run and IDs added above the starting limit constrain coverage.

### Inspect an existing local file

Existing local files can be inspected directly. This uses the same helper without a SharePoint connection and accepts an explicitly named file regardless of extension:

```powershell
$localScan = @{
    Path = 'D:\SPOInventory\Downloads\Example.doc'
    ResultPath = Join-Path $env:TEMP (
        'LegacyWordInspection_' + [guid]::NewGuid().ToString('N') + '.json'
    )
}

.\Inspect-LegacyWord.ps1 @localScan
```

Replace `Path` with an existing document. The result is JSON, not the main script's Excel report set. The helper refuses to overwrite an existing result path. The inventory normally supplies the hard child-process timeout; the direct helper also has internal parser time/bounds checks but is not an independently supervised process launcher.

## What was tested

The v3.10.1 patch passed **476 offline regression checks**. The path tests used separate .NET hosts with both compatibility switches enabled and disabled; no production host configuration was changed.

| Test suite | Passed checks |
| --- | ---: |
| Native scanner | 38 |
| Report and scheduler | 59 |
| Earlier format coverage | 48 |
| Expanded Word formats | 75 |
| Isolated format/report pipeline | 68 |
| Legacy .NET path compatibility | 94 |
| Modern .NET path compatibility | 94 |

The legacy host reproduced the exact original `FileIOPermission` exception. The corrected code then passed flat and mirrored downloads, twelve folder levels, 455-character paths, Unicode/space/#/% names, hash and length checks, overwrite prevention, junction rejection, cleanup and both scanner execution paths. Tests also exercised the production download function with local SharePoint-shaped test doubles. UNC path normalization was tested; real network-share I/O and every Windows Server release were not tested.

The additional corpus contained 16 pinned public samples—14 Apache POI documents and two oletools clean XML examples—and 13 actual Word SaveAs exports. The complete pipeline processed all 29: 25 complete inspections, four expected review results and no processing failures. Three encrypted documents and one HTML document with a separate Office sidecar required review. Five inspectable containers retained VBA indicators.

Samples were obtained from the pinned [Apache POI document fixtures](https://github.com/apache/poi/tree/07f1631dd66f647aacf2f43e3f408fe033aac4d0/test-data/document) and [oletools XML fixtures](https://github.com/decalage2/oletools/tree/ec10260989dbc48b9109e3d05d22beee19cf2333/tests/test-data/ooxml). The public corpus uses the pinned repository revisions linked here. Sample documents and fixture builders are development resources and are not included in the production ZIP. Word exports were produced with macros forced disabled, using plain input and a previously reviewed benign VBA test project. No VBA was executed.

**Historical v3.10 live validation:** the earlier QA library run enumerated 77 items: 60 files and 17 folders. It processed all 57 matching `.doc` files successfully, detected VBA indicators in 47, exported 513 metadata rows and recorded no processing failures. The run completed in approximately 91 seconds. Download hashes, correlation evidence, Excel row counts and paths reaching six folder levels were verified. That timing describes the earlier small QA run; it is not a production sizing estimate. The v3.10.1 patch was tested offline and has not had a new live tenant run.

The selected real binary files exercised effective revisions 193, 257 and 274. The scanner maps additional Word binary revisions, but testing did not include an authentic sample from every Word release since 1997. Original deleted v3.8 problem files were unavailable, so their exact byte-level behavior cannot be retrospectively certified.

During the earlier v3.10 fixture preparation, Microsoft Defender reported no threats in the fixture folder. That result does not guarantee that downloaded files are free of security issues. One development-only inline diagnostic command was blocked by AMSI; protections remained enabled, and the normal scanner/test scripts ran with protections enabled. The public validation summary records this result and its limitations; development fixtures are not part of the runtime download.

## Network activity and local retention

Network use and local retention should be understood before a large run. Normal operation communicates with SharePoint and Microsoft authentication services. Dependency acquisition uses PSGallery and package-provider infrastructure when installation is needed. The content parser itself has no network operations: it does not upload file bytes for third-party analysis or follow external templates, URLs or HTML sidecars.

Downloaded files and reports remain in the selected local directories. Reports can contain names, paths, user metadata and diagnostic information. Apply the organization's normal storage and retention requirements to those outputs. This behavior does not guarantee that endpoint protection or data-loss-prevention controls will never flag a download or PowerShell operation.

## Troubleshooting

Use the recorded evidence to resolve operational issues. Common cases have specific next steps:

| Symptom | What to check |
| --- | --- |
| `CreateDirectory` reports illegal characters / `FileIOPermission` | Replace all four runtime files with v3.10.1 and start a fresh PowerShell process; do not strip long-path prefixes manually |
| A different PnP version is already loaded | Start a fresh Windows PowerShell 5.1 STA session and use the pinned dependency |
| Sign-in fails or no usable MFA window appears | Check the chosen authentication mode and tenant policy; use `SystemBrowser` only with an approved app ID |
| HTTP 403 | Confirm the signed-in account's access to the target web/library/files and the relevant tenant restrictions |
| Library root cannot be found | Use the actual document-library root; provide `-SiteUrl` for a nested subweb |
| A folder is rejected | Use a decoded library-relative path such as `Policies/Finance`; in count mode check its depth against `MaxDepth` |
| Processed counts are temporarily static | Check active workers, retries and cooldown; batch results appear when workers complete |
| Many files show `Unknown` | Inspect Format, DetectionScope, ErrorStage and diagnostics; distinguish encryption, unsupported structures, limits and parser failures |
| A file has VBA evidence and a warning | Preserve the positive evidence; examine incomplete-source or version-correlation details |
| File type is `NotOLE` | Review its signature and format diagnostic; do not interpret it as macro-free |
| Content was not downloaded | Check the selected mode, permissions, `MaxScanMB`, disk space and DownloadStatus |
| A local file is difficult to locate | Follow the exact `LocalPath` in its document/review row |
| Some matching files remain unprocessed | Check EnumerationComplete, NotProcessed, RunErrors and the execution log |

Built-in command help is available from the deployed scripts:

```powershell
Get-Help .\Export-SPOLegacyDocInventory.ps1 -Examples
Get-Help .\Export-SPOLegacyDocInventory.ps1 -Full
```

The [deployment README](/scripts/legacy-doc-inventory/README.md), [format coverage matrix](/scripts/legacy-doc-inventory/FORMAT-COVERAGE-v3.10.md), [public validation summary](/scripts/legacy-doc-inventory/VALIDATION-v3.10.json) and [SHA256 checksums](/scripts/legacy-doc-inventory/SHA256SUMS.txt) accompany the release. A practical first run is a folder count, followed by a metadata inventory, then content inspection of the areas that need it.

[Download v3.10.1](/scripts/legacy-doc-inventory/Export-SPOLegacyDocInventory-v3.10.1.zip) · [Browse the four source files](https://github.com/pvutla001/pvutla001.github.io/tree/main/scripts/legacy-doc-inventory) · [Return to articles](/#articles)
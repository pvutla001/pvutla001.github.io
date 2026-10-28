# SharePoint legacy Word inventory v3.10.1

[Read the full guide](https://pvutla001.github.io/sharepoint/legacy-doc-inventory/).

## v3.10.1: fix for Windows PowerShell path errors

Version **3.10.1** fixes repeated download failures such as `Exception calling "CreateDirectory": Illegal characters in path`. The v3.10 engine passed a `\\?\` device path to a .NET Framework string-path API. Hosts using legacy path validation could reject that prefix in `FileIOPermission`, even for a short, valid destination. The failure occurred while preparing the local download directory, before Word format or VBA inspection.

The correction covers directory creation, file existence checks, download streams, partial-file renaming and cleanup, snapshot length/hash reads, mirrored-path checks, scanner input/output and path-identity comparison. Document I/O now uses Windows Unicode file APIs and handle-based .NET streams, preserving long paths without changing the registry or `powershell.exe.config`. Microsoft describes the [Framework compatibility behavior](https://learn.microsoft.com/en-us/dotnet/framework/migration-guide/mitigation-path-normalization), [Unicode directory API](https://learn.microsoft.com/en-us/windows/win32/api/fileapi/nf-fileapi-createdirectoryw) and [handle-based FileStream constructor](https://learn.microsoft.com/en-us/dotnet/api/system.io.filestream.-ctor?view=netframework-4.8).

**Upgrade all four runtime files together, then start a fresh Windows PowerShell 5.1 STA process.** Commands and dependencies remain the same. The summary identifies version `3.10.1`, and the log records the host's path settings without changing them. Windows file failures include the operation, path and Win32 error number. No Python, additional module, separate C# SDK or app registration is introduced by this fix.

Long-path support here applies to retained document files, including deep `Mirror` layouts. Keep the script, module and report roots short because the launcher and reporting code still use standard PowerShell/.NET path APIs; for example, use `C:\Tools\SPOInventory` and `-OutputDirectory 'C:\SPOReports'`. Windows/filesystem component limits and access controls still apply. The former `v3.10.zip` download address remains a compatibility alias for the corrected package.

## Required files

Keep these four matching files together. Execute the first file; it loads the other helpers as required.

1. Export-SPOLegacyDocInventory.ps1
2. SPOLegacyDocInventory.Engine.ps1
3. SPOLegacyDocInventory.Runtime.cs
4. Inspect-LegacyWord.ps1

README.md, FORMAT-COVERAGE-v3.10.md, VALIDATION-v3.10.json and SHA256SUMS.txt are optional reference files. Tests, fixture builders, sample documents and historical reports are not runtime dependencies.

## Prerequisites

- Windows PowerShell 5.1 STA, preferably 64-bit, and Windows .NET Framework.
- Permission to run approved scripts, compile the bundled helper with Add-Type and launch child PowerShell processes.
- Read/download access to the intended library and files; writable local output locations.
- PnP.PowerShell 1.12.0. The script saves this pinned module under .dependencies if missing. It can install the NuGet provider for the current user if needed. Use -NoInstall to disable installation.
- Network access to the SharePoint tenant and Microsoft authentication services. Dependency installation also needs package-repository access.

No Python, Word, Excel, ImportExcel, Visual Studio or separate C# SDK is required. The C# source is compiled through built-in Add-Type. Do not replace the pinned module with a current PnP version in this release.

## Quick start

Start a fresh session:

```powershell
powershell.exe -NoProfile -STA
```

In that window, change to the directory containing the four files:

```powershell
Set-Location 'C:\Tools\SPOLegacyDocInventory\V3.10'
$library = 'https://contoso.sharepoint.com/sites/Example/Shared%20Documents'

# Metadata only: no content downloads, all folder depths.
.\Export-SPOLegacyDocInventory.ps1 -LibraryUrl $library

# Counts by folder: every depth, no content downloads.
.\Export-SPOLegacyDocInventory.ps1 -LibraryUrl $library -CountOnly -MaxDepth -1

# Download and inspect VBA indicators, retain folder hierarchy and log every result.
.\Export-SPOLegacyDocInventory.ps1 -LibraryUrl $library -CheckBoth -DownloadLayout Mirror -Verbose
```

Choose the command for the mode you need. Complete the script's sign-in/MFA. Default LegacyWebLogin retains UseWebLogin; tenant policy can block it. SystemBrowser requires an approved interactive app ClientId and configuration. The script does not create an app or bypass tenant policy.

For restricted installation, pre-stage the complete PnP.PowerShell 1.12.0 module and dependencies in .dependencies beside the scripts or another PSModulePath location, then add -NoInstall. This disables installation, not tenant network access.

Use -Folder 'Policies/Finance' to select that library-relative folder and its descendants. Count mode defaults to MaxDepth 2; -1 means all depths. Metadata/content modes include all selected descendants. Discovery still pages library items even when Folder narrows processing.

## Reports and interpretation

Timestamped reports/logs are written beneath DocInventoryReports; retained downloads beneath DocInventoryDownloads. Both are created automatically. OutputDirectory and DownloadDirectory override the parent locations.

Open RunSummary first, then AllDocuments_Metadata, SuccessfulDocuments and ReviewAndFailures. Review rows map retained files through LocalPath; failed files are not duplicated into a separate folder. FailureDetails and DocumentDiagnostics retain detailed errors. Verbose also records successful documents.

Word97PlusOLE is a normal Word binary label. The scanner also recognizes supported RTF, Word XML, Flat OPC, MHTML/ActiveMime, HTML and renamed OpenXML contents. Library selection remains .doc filenames only. Count/metadata modes do not inspect bytes. Content inspection downloads files and never executes Word or VBA.

Unknown/Unsupported/Partial require review. Encrypted content and unresolved HTML sidecars are not macro-free findings. WordBasic, packed embedded objects, external templates and compiled-only/stomped VBA are not exhaustively assessed. NotDetected is not a safety verdict or a guarantee of Word Online compatibility.

476 offline regression checks passed for v3.10.1, including both .NET path-compatibility modes, 12 folder levels and 455-character paths. The 57-document live QA result belongs to the earlier v3.10 release; no new tenant run or 200,000-file benchmark was performed for this patch. See VALIDATION-v3.10.json for the current patch evidence and limitations.
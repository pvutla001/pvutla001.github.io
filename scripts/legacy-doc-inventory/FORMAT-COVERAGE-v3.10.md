# Word format coverage — v3.10

> v3.10.1 retains this format coverage and corrects local document path handling. See README.md and VALIDATION-v3.10.json for the patch tests.

These 29 additional samples passed the expected-result tests and actual isolated scanner/report pipeline: 25 complete inspections, four expected review records, no processing failures. Five containers retained a VBA project. VBA was never executed. All 366 regression checks passed. The separately tested live library had 57 matching documents and no processing errors.

| Sample | Actual format | Scan status | VBA indicators | Projects | Effective FIB |
| --- | --- | --- | --- | ---: | ---: |
| ApachePOI_capitalized.doc | Word97PlusOLE | Scanned | NotDetected | 0 | 274 |
| ApachePOI_documentProperties.doc | Word97PlusOLE | Scanned | NotDetected | 0 | 257 |
| ApachePOI_endingnote.doc | Word97PlusOLE | Scanned | NotDetected | 0 | 257 |
| ApachePOI_footnote.doc | Word97PlusOLE | Scanned | NotDetected | 0 | 257 |
| ApachePOI_HeaderFooterUnicode.doc | Word97PlusOLE | Scanned | NotDetected | 0 | 274 |
| ApachePOI_hyperlink.doc | Word97PlusOLE | Scanned | NotDetected | 0 | 257 |
| ApachePOI_Lists.doc | Word97PlusOLE | Scanned | NotDetected | 0 | 274 |
| ApachePOI_password_password_cryptoapi.doc | EncryptedOLE | Unknown | Unknown |  |  |
| ApachePOI_password_tika_binaryrc4.doc | EncryptedOLE | Unknown | Unknown |  |  |
| ApachePOI_PasswordProtected.doc | EncryptedOLE | Unknown | Unknown |  |  |
| ApachePOI_simple.doc | Word97PlusOLE | Scanned | NotDetected | 0 | 193 |
| ApachePOI_simple-table.doc | Word97PlusOLE | Scanned | NotDetected | 0 | 193 |
| ApachePOI_simple-table2.doc | Word97PlusOLE | Scanned | NotDetected | 0 | 274 |
| ApachePOI_table-merges.doc | Word97PlusOLE | Scanned | NotDetected | 0 | 257 |
| Oletools_harmless-clean.xml | WordFlatOPC | Scanned | NotDetected | 0 |  |
| Oletools_harmless-clean-2003.xml | Word2003XML | Scanned | NotDetected | 0 |  |
| Plain_FilteredHTML.doc | HTML | Scanned | NotDetected | 0 |  |
| Plain_FlatOPC.doc | WordFlatOPC | Scanned | NotDetected | 0 |  |
| Plain_HTML.doc | HTML | Scanned | NotDetected | 0 |  |
| Plain_MHTML.doc | MHTML | Scanned | NotDetected | 0 |  |
| Plain_StrictOpenXML.doc | WordOpenXML | Scanned | NotDetected | 0 |  |
| Plain_UnicodeText.doc | PlainText | Scanned | NotDetected | 0 |  |
| Plain_Word2003XML.doc | Word2003XML | Scanned | NotDetected | 0 |  |
| VBA_BinaryTemplate.doc | Word97PlusOLE | Scanned | Detected | 1 | 274 |
| VBA_FlatOPCMacro.doc | WordFlatOPC | Scanned | Detected | 1 |  |
| VBA_HTML.doc | HTML | Unknown | Unknown |  |  |
| VBA_MHTML.doc | MHTML | Scanned | Detected | 1 |  |
| VBA_OpenXMLTemplate.doc | WordOpenXML | Scanned | Detected | 1 |  |
| VBA_Word2003XML.doc | Word2003XML | Scanned | Detected | 1 |  |

The HTML positive source uses a separate editdata.mso file; the HTML record intentionally remains Unknown because the scanner does not follow sidecars or URLs. Encrypted files are not decrypted. Public samples come from pinned Apache POI and oletools commits; the full guide links the pinned source repositories. Fixture files and development tools are not included in the production ZIP. Word exports were generated using Word SaveAs with macros forced disabled, from plain text and a previously reviewed benign test project.

Additional malformed cases cover truncated MIME boundaries, invalid base64, corrupted ActiveMime, DTDs, missing/renamed VBA parts, unresolved references, size limits and invalid paths. A good document after an error is tested. This is static VBA-container coverage, not a complete Office security verdict, exact application-version detector or Word Online compatibility test.

The inventory discovers .doc filenames. XML/MHTML/OpenXML examples saved under that extension are parsed by their bytes; genuine .docx/.docm/.xml names are not added to library discovery. The direct helper can inspect explicitly named local paths with other extensions.

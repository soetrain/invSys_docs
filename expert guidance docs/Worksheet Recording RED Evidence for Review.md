# Worksheet recording: sanitized historical RED evidence

Prepared 2026-11-11 for review of Slice 4be-A under Architecture v4.11, Plan 022 and Controls 1.735. This is a read-only extraction of the latest completed 2026-11-05 `WorksheetRecording` diagnostic. It contains check identities and Boolean results only. No workbook content, operator input, credentials, activity payload, event identity, warehouse identity or machine/session path is included.

## Provenance and limit

- The source `red.json` has SHA256 `3A54BC5B73E48FAB677F74F3EA5BE86B68AAD5FBF4A042528280659A8886E7DE`. Its 78 ordered checks are **50 PASS / 28 FAIL**. The five latest November 5 copies are byte-identical with the same ordered results.
- The diagnostic used the older `validation-inventory-uom-02` five-XLAM candidate. Its Core package hash begins `959D72ED`; PR4-BASE-03 Core begins `57273303` and PR4-BASE-05 Core begins `970D829B`. The diagnostic worker hash also differs from the present worker. The old run closed naturally with settings and packages preserved; worker exit 1 is the expected RED result. It was not an acceptance run.
- The current PR4/PR11 source already lists both `ADMIN_INVENTORY_CREATE_TABLE` and `ADMIN_INVENTORY_UPLOAD_TABLE` in catalog version 47 and has `modInventoryWorksheetActivity.Outcome`. Their behavior on the current packages **has not been checked by this historical run**. Requalify the focused worksheet selector on a current pinned candidate before treating any of the 28 failures as current.

## Exact failed identities

Catalog and declared outcome checks: **14**. `Catalog47` checks require a definition with the specified owner, capability and command class. `Conclusion` checks require the specified outcome/data effect and completion evaluation. The Boolean file does not disclose which subcondition failed.

```text
WorksheetRecording.Catalog47.CREATE_TABLE
WorksheetRecording.Conclusion.CREATE_TABLE.REQUESTED
WorksheetRecording.Conclusion.CREATE_TABLE.SAVED
WorksheetRecording.Conclusion.CREATE_TABLE.REFUSED
WorksheetRecording.Conclusion.CREATE_TABLE.PARTIAL
WorksheetRecording.Conclusion.CREATE_TABLE.FAILED
WorksheetRecording.Catalog47.UPLOAD_TABLE
WorksheetRecording.Conclusion.UPLOAD_TABLE.REQUESTED
WorksheetRecording.Conclusion.UPLOAD_TABLE.CONFIRMED
WorksheetRecording.Conclusion.UPLOAD_TABLE.EMPTY
WorksheetRecording.Conclusion.UPLOAD_TABLE.REJECTED
WorksheetRecording.Conclusion.UPLOAD_TABLE.REFUSED
WorksheetRecording.Conclusion.UPLOAD_TABLE.PARTIAL
WorksheetRecording.Conclusion.UPLOAD_TABLE.FAILED
```

Upload observation and retry checks: **11**. Each exact-record check expects a paired request and terminal observation when recording is enabled. The correlation checks require one attempt and one owner-derived outcome under the same activity identity. Source-reference checks require the exact owner-returned event identity and submission state. The historical Boolean report does not provide actual row counts or observed values.

```text
WorksheetRecording.UploadEnabled.ExactRecordCount
WorksheetRecording.UploadEnabled.ExactCorrelatedOwnerOutcome
WorksheetRecording.UploadEnabled.ExactOwnerEventReferences
WorksheetRecording.UploadSaveThrows.ExactRecordCount
WorksheetRecording.UploadSaveThrows.ExactCorrelatedOwnerOutcome
WorksheetRecording.UploadSaveThrows.ExactOwnerEventReferences
WorksheetRecording.UploadSaveThrows.Retry.CorrelatedConfirmation
WorksheetRecording.UploadSaveThrows.Retry.ExactPriorEvent
WorksheetRecording.UploadQueueThrows.ExactRecordCount
WorksheetRecording.UploadQueueThrows.ExactCorrelatedOwnerOutcome
WorksheetRecording.UploadQueueThrows.ExactOwnerEventReferences
```

Config preservation checks: **3**. Each checks that Config bytes stayed equal or that the scoped diagnostic established a permitted save-metadata cause. These failures require separate D8-INV-CONFIG-01 assessment; the Boolean report does not identify changed binary parts or authorize relaxing the check.

```text
WorksheetRecording.UploadEnabled.ConfigBytesOrExplainedSaveMetadata
WorksheetRecording.UploadSaveThrows.ConfigBytesOrExplainedSaveMetadata
WorksheetRecording.UploadQueueThrows.ConfigBytesOrExplainedSaveMetadata
```

## Passing boundaries relevant to diagnosis

The same run passed all five instrumented package compiles and the actual bound handler for each of the three Upload cases. Independent canonical outcomes, saved operator results, unknown-column and unrelated-content preservation, redaction, Auth byte preservation and read-only Config extraction passed in those cases. The injected queue-write boundary was reached; the failed observation checks must not mask those passing business checks. Earlier Create cases and other selectors are outside this 78-check diagnostic.

## Review pointers and next step

Use the current `tests/tooling/Slice4beInventoryWorksheetRecording.ps1` for exact assertions, especially `Test-InventoryWorksheetRecording` and its catalog/Upload branches; `src/Core/Modules/modActivityCatalog.bas`, `modInventoryWorksheetActivity.bas`, `modActivityReferences.bas` and the actual Admin worksheet handlers for source mapping. Architecture v4.11 governs outcome semantics, source references, policy, redaction and authority. First re-establish a focused packaged RED on the current pinned candidate through the real handler; a missing fixture, compile error, adapter failure or old-package RED is not a valid current behavioral RED. Then separate any remaining observation failure from Config serialization evidence before proposing a product change.

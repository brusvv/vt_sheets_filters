# Visible Sort Transfer

## Status and rollout gate

The transfer implementation is present, but it is **not enabled by default**.

`CONFIG.visibleSortTransferEnabled` is `false` in the repository. Keep it false in the working deployment until the live integration procedure below passes on a disposable copy of the workbook. Manual Sort and AutoSort both use the same detailed sorting path, so changing this flag enables the transfer for both paths.

The mechanism requires the Apps Script **Advanced Sheets service**. In the Apps Script editor, open **Services → Add a service → Google Sheets API**. If the script is attached to a standard Google Cloud project, also confirm that the Google Sheets API is enabled in that Cloud project. This repository does not contain an `appsscript.json`, so the service is not silently added to a guessed manifest.

## Transfer algorithm

The existing Resolver and Planner remain the sources of visibility and row ordering. The transfer module consumes their output; it does not infer why a row is hidden, change FilterState or ManualState, or reintroduce FIFO key matching.

1. Validate the plan against the visibility mask. Source and destination counts must match, every source and destination must be a visible physical row, and each block must remain a permutation of its own visible slots. Any filter-hidden row or merged cell intersecting C:L blocks the transfer.
2. Capture a full C:L cell snapshot. It includes values, R1C1 formulas, notes, rich-text content and runs, data-validation descriptors, and the cell-format properties exposed by Apps Script.
3. Create two uniquely named, hidden temporary sheets: `__VS_SRC_*` and `__VS_OUT_*`. The original C:L range is copied to the source stage at the same physical coordinates. Every source row therefore exists in staging before a destination is written.
4. Use Sheets API `CopyPasteRequest` requests to construct the proposed result on the output stage. Each copy operation includes `PASTE_NORMAL` and `PASTE_DATA_VALIDATION`. Adjacent contiguous mappings are grouped to reduce request count.
5. Compare each staged destination against its staged source. Formula comparison uses R1C1 notation so relative formulas can move to their destination row without being mistaken for formula changes. Values, notes, rich text, validation and supported format properties are also compared.
6. Recheck the source snapshot and call `validateSortVisibilityContextFresh_()` immediately before the destination write. A stale data, FilterState, ManualState, sheet, range or physical-visibility snapshot aborts before target cells are written.
7. Persist a replay guard, then send one `spreadsheets.batchUpdate` containing all destination copy requests. The requests target only the planner's visible physical destination rows in C:L. Hidden and non-participating physical rows are not write targets.
8. Read the complete original C:L range back and compare it with the expected post-sort snapshot. The expected snapshot moves only planned visible rows; it requires hidden rows and every non-participating position to remain byte-for-byte equivalent across the tracked cell properties.
9. Delete temporary sheets only after successful verification.

The official Sheets API documentation describes `PASTE_NORMAL` as copying values, formulas, formats and merges; data validation has its own `PASTE_DATA_VALIDATION` type. Notes and rich-text runs are not assumed to be preserved based on the paste-type description: the implementation snapshots and verifies them. If a live API round-trip does not preserve a tracked property, the pre-write staging verification blocks the commit. The Apps Script snapshot checks the formatting properties listed above, but does not independently serialize every border/padding detail or conditional-format rule metadata; those must be explicitly inspected during the live copy-workbook test. Conditional-formatting rules are sheet-level rules and are not separately migrated by this cell-range transfer.

The batch-update API validates requests before applying them and applies the request batch atomically when accepted. That does not prevent collaborators or UI edits from interacting with the sheet around the request. ScriptLock serializes cooperating script executions only; it does not lock out spreadsheet users.

## Error handling and recovery

| Status | Meaning | Target C:L writes |
|---|---|---|
| `TRANSFER_NOT_VERIFIED` | Feature flag/service unavailable or plan rejected | None |
| `BLOCKED_STALE_CONTEXT` | Data/state/visibility changed before commit | None |
| `BLOCKED_STAGING` | Stage creation/copy/snapshot failed | None |
| `SUCCESS` | Transfer and full-range verification succeeded | Verified |
| `TRANSFER_WRITE_UNKNOWN` | Destination batch threw after the write attempt began | Outcome may be unknown |
| `TRANSFER_VERIFY_FAILED` | Destination write returned, but post-write verification failed | May already be changed |

There is deliberately no automatic rollback after the destination write attempt. A blind rollback could overwrite legitimate user edits made after or during the operation. On `TRANSFER_WRITE_UNKNOWN` or `TRANSFER_VERIFY_FAILED`, the two hidden staging tabs are retained as recovery evidence. Do not rerun the same operation: the document-property replay ledger blocks that operation ID. Inspect the target and both `__VS_*` tabs manually, copy back only after confirming the correct recovery point, and remove the tabs only after the incident is resolved.

The replay ledger keeps the 40 most recent operation IDs in Document Properties. A failure before destination writes is marked as aborted-before-write and may be retried; once a destination write is attempted, the same operation ID is blocked.

Merged cells intersecting C:L are rejected. The implementation does not move whole sheet rows: cells outside C:L, row heights, row groups, filters, and physical row visibility remain attached to their original physical positions. Conditional-formatting rules remain sheet-level rules and are not rewritten by the transfer.

## Node tests

Run the repository suite:

```sh
node --test tests/*.test.cjs
```

`tests/visibleSort.transfer.test.cjs` exercises the transfer module with a mock Sheets API. It covers ordinary permutations, manually hidden slots, filter-hidden fail-closed behavior, alternating and consecutive hidden groups, zero/one visible rows, stale context, visibility changes between planning and write, count mismatch, staging failure, destination-write failure, verification failure without rollback, replay prevention, and preservation checks for values, R1C1 formulas, formatting, notes, rich text, data validation, hidden slots and cells outside C:L.

These tests verify code paths and the transfer contract in a mock environment. They do **not** prove live Google Sheets API behavior.

## Live integration test on a disposable spreadsheet

1. Make a separate copy of the workbook. Do not use the production workbook for the first run.
2. In the copied Apps Script project, enable **Google Sheets API** under **Services**. Enable the Cloud API too if the project uses a standard Cloud project.
3. In the copy only, change `CONFIG.visibleSortTransferEnabled` to `true`. Keep the production setting false.
4. Use a test sheet with a dedicated C:L data block. Include at least two scenario blocks and deliberately scrambled scores/Dots. Include ordinary values, formulas with relative and absolute references, a rich-text hyperlink, a note, a data-validation rule, non-default number formats, backgrounds, font styles and alignment.
5. Create a mix of visible and manually hidden rows: alternating hidden/visible rows and at least two consecutive hidden-row groups. Record the hidden rows' physical row numbers and their C:L contents. Keep at least one visible row and one hidden row outside the changed scenario blocks.
6. Ensure there are no merged cells intersecting C:L; this implementation intentionally rejects them.
7. Record C:L, columns A:B and M onward, row heights, hidden-row positions, FilterState and ManualState before the test.
8. Run each Manual Sort mode on the copy. Check the toast, execution log and operation trace. Verify that only visible C:L slots moved in each scenario block, hidden rows remain at their original physical positions, and all non-participating rows and columns remain unchanged.
9. Compare formulas using `getFormulasR1C1()`. Check values, rich-text links/runs, notes, data validations, number formats, backgrounds, fonts, alignments and wrapping. Confirm both temporary tabs are deleted after `SUCCESS`.
10. Repeat with AutoSort enabled on the copy and run the existing filter/reset workflows. Confirm FilterState and ManualState are unchanged by the transfer and that physical visibility is still correct.
11. Separately verify that a Basic Filter / filter-hidden state is rejected before any destination write, and that an injected stale context is also rejected. Do not assume a filter-hidden context is sortable: the existing Resolver deliberately fails closed for it.
12. Save the execution log and before/after evidence. Only after the live tests pass should the owner consider enabling the feature in the intended deployment.

Stop and do not retry after `TRANSFER_WRITE_UNKNOWN` or `TRANSFER_VERIFY_FAILED`. Preserve the hidden staging tabs and inspect them before any manual recovery.

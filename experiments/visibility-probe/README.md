# Visibility Probe v1.3.0

Standalone diagnostic tool for comparing physical row visibility from Apps Script and the Advanced Google Sheets API. This experiment is isolated from production code and must be run only in a disposable spreadsheet / separate Apps Script project.

## What it measures

For the same configured physical row ranges, the probe reads these two independent flags using both methods:

- `hiddenByUser` via `Sheet.isRowHiddenByUser(rowNumber)` and `GridData.rowMetadata[].hiddenByUser`.
- `hiddenByFilter` via `Sheet.isRowHiddenByFilter(rowNumber)` and `GridData.rowMetadata[].hiddenByFilter`.

A present `rowMetadata` object with an omitted Boolean property is interpreted as the protobuf JSON default `false`. This is distinct from missing/invalid row metadata: absent row metadata objects, invalid Boolean types, missing rows, duplicate row metadata, invalid `GridData.startRow`, out-of-range metadata and API errors remain incomplete or errors. API mapping uses `sheetId` and each `GridData.startRow`; it never uses response-array order.

The tool is read-only: it does not edit cells, visibility, filters, sheet state, FilterState, ManualState, project settings, or triggers. It does not read cell values. It writes one structured JSON report to the Apps Script execution log and returns the same report.

## Repository layout

- `visibilityProbe.gs` — source to copy into a separate Apps Script project.
- `appsscript.standalone.json` — manifest variant for a standalone project.
- `appsscript.bound.json` — manifest variant for a container-bound project.
- `../../tests/visibilityProbe.test.cjs` — Node mock tests for parser, comparator, report and benchmark behavior.

Do not copy this file into the production Apps Script project. The source lives under `experiments/visibility-probe/` specifically to keep it separate from the main sorting/filtering implementation.

## Choose the project type and permissions

### Standalone Apps Script project

The standalone variant uses `SpreadsheetApp.openById()`. The manifest requests `https://www.googleapis.com/auth/spreadsheets`; this scope technically permits spreadsheet writes even though the probe source has no write calls. Use a disposable spreadsheet and a separate script project. Do not add this scope to the production project.

1. Create a disposable copy of the workbook.
2. Create a new standalone Apps Script project at <https://script.google.com/>.
3. Copy only `visibilityProbe.gs` into that project.
4. In Project Settings, enable `appsscript.json` and copy the contents of `appsscript.standalone.json` into it (the manifest must be named exactly `appsscript.json`).
5. In **Services**, add/verify **Google Sheets API** (Advanced Sheets service v4). The manifest declares this service as `Sheets`.
6. In the associated Google Cloud project, verify that Google Sheets API is enabled.
7. Set `spreadsheetId` and either `sheetId` or `sheetName` in `VISIBILITY_PROBE_CONFIG`. `sheetId` takes precedence. The `sheetId` is the `gid` query value in a sheet URL.
8. Set the row ranges and scenario label, save, then run `runVisibilityProbe()`.

### Container-bound Apps Script project

This option uses narrower scopes, but the project must be created from the disposable spreadsheet itself.

1. Open the disposable copy and choose **Extensions → Apps Script**.
2. Copy only `visibilityProbe.gs` into the bound project.
3. Use the contents of `appsscript.bound.json` as the project's `appsscript.json`.
4. Enable the Google Sheets API advanced service under **Services**, and enable the API in the associated Cloud project if required.
5. Leave `spreadsheetId`, `sheetId`, and `sheetName` blank to use the active spreadsheet and active sheet. Select the intended test sheet before running.

The bound manifest combines `spreadsheets.currentonly` for the active spreadsheet with `spreadsheets.readonly` for the Advanced Sheets API read. Neither manifest belongs in the production Apps Script project.

## Configuration

The default range is intentionally small (rows 3–22). For the baseline Manual Sort block from the 14 November 2026 log, set the range to rows 3–131 inclusive:

```javascript
ranges: [
  { startRow: 3, numRows: 129 }
],
scenarioLabel: 'visible-baseline',
benchmark: {
  warmupRuns: 2,
  sampleRuns: 12
}
```

`ranges` accepts multiple non-overlapping `{ startRow, numRows }` entries. Rows use Google Sheets' 1-based physical row numbering. Overlapping, invalid, zero-length, or out-of-sheet ranges are rejected before measurement.

## Functions

### `runVisibilityProbeApiDiagnostics()`

Runs only the Advanced Sheets API read and logs one compact line prefixed with
`VISIBILITY_PROBE_API_DIAGNOSTICS`. The output includes the actual API exception
when the request fails, or the returned sheet/GridData/rowMetadata structure
when it succeeds. It does not print row-by-row Apps Script results, so it avoids
the large execution-log truncation that can hide the `errors` array in the full
probe report.

Run this function first if the API request is failing or returns zero GridData.
Copy the single diagnostic line from the execution log.

### `runVisibilityProbe()`

Runs one comparison for the selected scenario. It reads all selected rows through Apps Script, makes one `Sheets.Spreadsheets.get()` call for all configured A1 ranges, maps the returned metadata, compares each flag independently, and returns/logs a JSON object.

This is a diagnostic comparison, not a performance benchmark. Its `performance.totalElapsedMs` includes target setup, both reads, mapping, comparison and report preparation.

### `runVisibilityProbeBenchmark()`

Runs the same comparison repeatedly, with defaults:

- 2 warm-up rounds per method;
- 12 measured rounds per method;
- alternating order: Apps Script first on even rounds, Sheets API first on odd rounds;
- one first invocation of each method separately reported as `performance.coldStart`;
- p50 and p95 using linear interpolation (Type 7 / `PERCENTILE.INC`).

At least one warm-up round is required so that the first invocation of each method remains outside the measured percentiles. `coldStart` means the first invocation of that method within the benchmark's Apps Script execution; it is **not** a measurement of Apps Script runtime startup across separate executions.

The benchmark reports three timing layers:

- `performance.appsScript.p50Ms` / `p95Ms`: successful complete Apps Script method passes.
- `performance.sheetsApi.p50RequestMs` / `p95RequestMs`, `p50ParseMs` / `p95ParseMs`, and `p50TotalMs` / `p95TotalMs`: only complete, type-valid Sheets API responses. API errors and incomplete responses are counted separately and excluded from these percentiles.
- `performance.pairedProbe.p50TotalMs` / `p95TotalMs`: wall time for both methods in one measured round plus row comparison. This includes failed or incomplete measured rounds so that it reflects the end-to-end diagnostic pair.

Each measured round includes successfully compared rows, rows matching both flags, mismatching rows, mismatches per flag, incomplete rows, and individual timings. `comparisonTotals` adds row observations across rounds; it counts repeated observations of the same physical rows, not unique sheet rows.

## Manual scenarios

The probe never creates or changes any scenario. Prepare each state manually in the disposable workbook, set `scenarioLabel`, then run `runVisibilityProbe()` and save that JSON report before changing to the next state.

1. `visible-baseline` — all selected rows visible.
2. `user-hidden` — manually hide selected rows with the row context menu.
3. `basic-filter-hidden` — apply a normal Basic Filter and select criteria that exclude rows.
4. `filter-view-active` — activate a Filter View that hides selected rows.
5. `collapsed-row-group` — group rows and collapse the group.
6. `user-and-basic-filter` — use rows that are both manually hidden and excluded by a Basic Filter, where the UI permits that combination.

Use the exact same row ranges for all six scenarios. Save a separate JSON line for every scenario. Do not assume the flags are equivalent across Filter Views: Apps Script `isRowHiddenByFilter()` documentation specifically excludes rows hidden by Filter Views. Record any difference between Apps Script and API as an experimental result; do not reinterpret it as a match.

## API request and row mapping

The probe makes one Advanced Sheets API request per API sample:

- `ranges`: all configured A1 row ranges in column A (cell values are not requested).
- `fields`: `sheets(properties(sheetId),data(startRow,rowMetadata(hiddenByUser,hiddenByFilter)))`. The field mask explicitly requests the GridData fields; `includeGridData` is omitted because the API ignores it whenever a field mask is present.

The `fields` mask is what requests the selected `GridData` fields. No `CellData` or cell values are requested. Each returned sheet is selected by `properties.sheetId`. Within it, each `GridData` maps `rowMetadata[index]` to physical row `startRow + index + 1`, since `startRow` is zero-based. Multiple `GridData` segments can contribute different parts of the configured ranges.

For a valid row metadata object, an omitted `hiddenByUser` or `hiddenByFilter` property is interpreted as `false`, matching the default-value omission behavior of protobuf JSON serialization. Explicit `true` and `false` values are preserved. This inference is limited to the two requested Boolean fields on a present metadata object; it does not apply when the metadata object or row is missing, duplicated, or malformed. Metadata outside the requested physical row set is recorded as an error and makes the API result incomplete. `rowMetadata` array length is reported as returned; it is not padded to the requested row count.

The `responseJsonUtf8Bytes` value is the UTF-8 size of `JSON.stringify()` over the Advanced Service response object. It is not a claim about exact HTTP wire bytes.

## Report fields

Both modes include:

- `probeVersion`, `generatedAt`, `mode`, `scenarioLabel`;
- `spreadsheetId`, `sheetId`, `sheetName`, `requestedRanges`;
- `status`, `apiStatus`, `apiGridDataCount`, `apiRowMetadataCount`;
- `comparisonSummary` with total, complete, matched-both-flags, mismatching and incomplete row counts, separate matches/mismatches for each flag, missing values and Apps Script read errors;
- `performance` and `errors`.

Single-run mode additionally includes the per-row comparison. Benchmark mode includes cold-start values, p50/p95, request counts, response size, paired comparison totals, and per-measured-round detail.

### Illustrative JSON shape (schema only; values are placeholders)

```json
{
  "probeVersion": "1.2.0",
  "generatedAt": null,
  "mode": "SINGLE_RUN",
  "status": "NOT_RUN",
  "scenarioLabel": "<scenario-label>",
  "spreadsheetId": null,
  "sheetId": null,
  "sheetName": null,
  "requestedRanges": [],
  "apiStatus": "NOT_RUN",
  "apiGridDataCount": 0,
  "apiRowMetadataCount": 0,
  "rows": [
    {
      "rowNumber": 3,
      "appsScript": { "hiddenByUser": null, "hiddenByFilter": null },
      "sheetsApi": { "hiddenByUser": null, "hiddenByFilter": null },
      "comparison": { "hiddenByUser": "UNAVAILABLE_API", "hiddenByFilter": "UNAVAILABLE_API" },
      "rowMetadataStatus": "MISSING_ROW_METADATA",
      "status": "INCOMPLETE"
    }
  ],
  "comparisonSummary": {
    "totalRowCount": 0,
    "completeRowCount": 0,
    "matchedRowsBothFields": 0,
    "mismatchedRows": 0,
    "incompleteRowCount": 0
  },
  "performance": { "totalElapsedMs": null },
  "errors": []
}
```

The example is only a schema illustration and is not a live result.

## Interpreting mismatches

- `MATCH` means both APIs resolve to the same Boolean for that physical row. An omitted API Boolean on a present `rowMetadata` object resolves to the default `false`.
- `MISMATCH` means both values are present but differ. Keep the two flags independent; a difference in one does not imply a difference in the other.
- `UNAVAILABLE_API`, `UNAVAILABLE_APPS_SCRIPT`, `INVALID_TYPE`, `MISSING_ROW_METADATA`, `DUPLICATE_ROW_METADATA`, and `INCOMPLETE` are not matches and must not be treated as false.
- `apiStatus: SUCCESS` means all requested physical rows had exactly one metadata candidate with both Boolean fields. It does not assert semantic equivalence to Apps Script.
- A Filter View mismatch is an expected possibility that must be captured, not normalized away.

A live equivalence/performance claim requires the six scenario reports and a repeated benchmark from the actual disposable Google Spreadsheet. Node tests validate only this probe's mapping and statistics logic; they do not establish real Google Sheets behavior or speed.

## Running Node tests

From the repository root:

```sh
node --test tests/*.test.cjs
```

The tests use mocked Apps Script and Sheets API objects. They do not require credentials or a live spreadsheet.
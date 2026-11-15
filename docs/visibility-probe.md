# Visibility Probe

An isolated, read-only experiment for comparing Apps Script row-visibility methods with `Sheets.Spreadsheets.get()` metadata.

The source, standalone/container-bound manifests, JSON output interpretation, six manual test scenarios, and full setup steps are documented in [experiments/visibility-probe/README.md](../experiments/visibility-probe/README.md).

The probe is not imported by production Resolver or sorting modules. It must be copied manually into a separate Apps Script project attached to a disposable spreadsheet. It does not change the production Resolver, Planner, Transfer, FilterState, ManualState, Tracer, or the Transfer feature flag.

Run the Node mock suite from the repository root:

```sh
node --test tests/*.test.cjs
```

These tests validate mapping and reporting logic only. They are not a live Google Sheets benchmark and do not establish semantic equivalence or performance of the Advanced Sheets API.

## Resolver batch-read rollout

The Resolver can read both physical visibility flags in one Advanced Sheets API
request through `Sheets.Spreadsheets.get()` and `GridData.rowMetadata`.
The implementation is in `sorting/sortVisibility.gs` and is gated by
`CONFIG.resolverBatchVisibilityEnabled`.

The flag defaults to `false`. When disabled, the Resolver continues to use
`isRowHiddenByUser()` and `isRowHiddenByFilter()` exactly as before. When
enabled, the Resolver uses one metadata request for the full scenario range.
The existing validation remains unchanged: any row with
`hiddenByFilter === true` is rejected as `EXTERNAL_FILTER_UNSUPPORTED`,
and FilterState/ManualState reconciliation still runs against the returned
physical visibility arrays.

The API reader treats an omitted Boolean property on a present
`rowMetadata` object as the protobuf default `false`. It does not treat a
missing rowMetadata object, missing row, duplicate sheet mapping, misaligned
GridData range, or invalid Boolean type as false; those conditions throw and
the Resolver fails closed.

The initial live probe established equivalence for visible rows, manually
hidden rows, and Basic Filter-hidden rows over 129 rows. The batch-read flag
remains off until the new Resolver path is exercised on a disposable copy of
the workbook and its fail-closed behavior is checked. The unit tests for the
metadata reader are not a substitute for that live Resolver test.

To perform the controlled live test, copy the updated source files into the
bound Apps Script project, set `resolverBatchVisibilityEnabled: true` in
`config/config.gs`, and run the existing sort operation only on a disposable
workbook. Confirm the Performance Tracer reports
`source: SHEETS_API_ROW_METADATA`, one Sheets API request, and zero calls to
both per-row Apps Script visibility methods. Restore the flag to `false`
unless the live test is intentionally being continued.

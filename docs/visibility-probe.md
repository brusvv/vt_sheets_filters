# Visibility Probe

An isolated, read-only experiment for comparing Apps Script row-visibility methods with `Sheets.Spreadsheets.get()` metadata.

The source, standalone/container-bound manifests, JSON output interpretation, six manual test scenarios, and full setup steps are documented in [experiments/visibility-probe/README.md](../experiments/visibility-probe/README.md).

The probe is not imported by production Resolver or sorting modules. It must be copied manually into a separate Apps Script project attached to a disposable spreadsheet. It does not change the production Resolver, Planner, Transfer, FilterState, ManualState, Tracer, or the Transfer feature flag.

Run the Node mock suite from the repository root:

```sh
node --test tests/*.test.cjs
```

These tests validate mapping and reporting logic only. They are not a live Google Sheets benchmark and do not establish semantic equivalence or performance of the Advanced Sheets API.

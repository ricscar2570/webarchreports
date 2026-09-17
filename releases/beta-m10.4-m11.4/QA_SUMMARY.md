# M10.4 / M11.4 — QA summary

Date: 2026-09-16
Status: SYSTEM-INTEGRATION GATE CANDIDATE; RC1 not frozen.

## Combined project evidence

- Production components: 77
- Recognized procedures: 1590
- Full-project checks: 1082/1082 PASS
- Existing components byte-identical vs M10.3 base: 70
- Shared-state contracts: 18/18 PASS

## Self-contained package evidence

- VERIFY_RELEASE: 67/67 PASS
- M10.4 model tests: 59/59 PASS
- Functional-flow models: 31/31 PASS
- System-integration models: 13/13 PASS
- Random integration: 5000 iterations / 20000 checks PASS
- System-integration mutations: 7/7 PASS
- Shared-state package audit: 6/6 PASS
- Runner: 7/7 suites PASS

## Archive evidence

Each of the three final ZIPs was reopened and verified after extraction. Internal manifest verification: 42/42 PASS for M10.4, M11.4 and the combined delivery.

## System-level invariant

No view that writes shared state may apply values if the state from which that view was built has changed in the meantime. M3 and M4 enforce this against both applied M2 values and PENDING drafts.

## Native gate still required

The following are not claimed PASS here: Microsoft VBE compile, Excel Workbook events, ListObject runtime behavior, Shape/OnAction execution, Save/close/reopen, ExportAsFixedFormat and performance on real datasets.

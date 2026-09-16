# WebArch Beta M10.4 / M11.4 — System Integration Hardening

Status: **SYSTEM-INTEGRATION GATE CANDIDATE — RC1 not frozen**.

M10.4 supersedes the previous M10.2 registry and the intermediate M10.3 hardening work. The release evaluates WebArch as one program rather than as isolated milestone modules.

The key system-level correction is concurrency protection between M3 `OP_PROTOCOLLI` and M4 `OP_PROTOCOLLI_PLUS`, which both edit M2 `PROTOCOLLO` state. Both views now fingerprint the applied M2 values and `PENDING` drafts captured at refresh time and reject Apply if another view or operation changed that shared state in the meantime. This prevents silent lost-update sequences M3→M4 and M4→M3 without coupling the two views directly.

The package also preserves the earlier M4/M5/M7/M8 safety hardening and M10 fail-closed source/state, discard, staleness and reporting gates.

Offline verification:

- Full project audit: 1082/1082 PASS
- Package release verifier: 67/67 PASS
- M10 model tests: 59/59 PASS
- Functional-flow models: 31/31 PASS
- System-integration models: 13/13 PASS
- Random integration: 5000 iterations / 20000 checks PASS
- Integration mutation tests: 7/7 PASS
- Full shared-state audit: 18/18 PASS
- Package runner: 7/7 suites PASS
- Extracted ZIP manifests: 42/42 PASS for M10.4, M11.4 and combined delivery

Microsoft VBE compile and real Excel runtime are **NOT EXECUTED** for this candidate. VBAProject compilation, Workbook events, ListObject runtime, Shape/OnAction, Save/reopen, PDF export and performance remain mandatory before RC1.

Exact binary/source delivery is archived in the WebArch Project Library. This repository path records status, system-integration evidence and checksums.

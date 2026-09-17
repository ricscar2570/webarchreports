# WebArch — RC Promotion Gate

## Current checkpoint

- Active branch: `beta`
- Candidate line: **M10.5 / M11.5 — STEP 10A**
- Stable accepted baseline: **EASYONE MINIMAL 1.0.3 + HOME 1.1**
- Current status: **OFFLINE FUNCTIONAL AUDIT + HOSTILE VBA SCAN PASS / RC BLOCKED**
- STEP10A package (Project Library): `WebArch_M10_5_STEP10A_VERSIONING_NAMING.zip`
- Package SHA-256: `31d68fb7eb6388f5d44c12ce5b79a51f852a9823323ad88671a5e26f58dc03fe`

`main` must not be changed until every blocking native gate below is closed. The accepted Alpha must remain recoverable and must not be rewritten by Beta remediation work.

## Offline evidence closed

The cumulative Beta remediation line currently records:

- STEP9 independent re-audit: **320/320 PASS**;
- STEP10A suite: **19/19 PASS**;
- STEP10A specific gate: **24/24 PASS**;
- STEP10A full static audit: **643/643 PASS**;
- source modules: **91**;
- STEP10A modified modules: **7**, limited to versioning/naming normalization;
- hostile VBA scan 17 September 2026: **PASS**;
- new confirmed P0 from hostile scan: **0**;
- new confirmed P1 from hostile scan: **0**;
- new confirmed P2 from hostile scan: **0**;
- confirmed STEP10A regressions from hostile scan: **0**.

Closed/remediated offline findings include writer/recovery/bootstrap, local snapshot semantics, procedure-scoped verification, M11 fail-safe, workbook-qualified M10 button bindings, EASYONE reentrancy guards, PDF provenance/freshness, runtime M10 version naming and M4 hardening naming consistency.

The previously known frozen-Alpha `modBacklog` `IIf` concern remains **LOW / LATENT_NOT_REPRODUCED** and is not a reason to modify the accepted Alpha.

## Open non-native items

### SURFACE-P1-01 — PENDING_REAL_WORKBOOK_GATE

Macro surface is classified offline, but no candidate may be privatized/removed until the real operational `.xlsm` is audited for persisted `Shape.OnAction` bindings.

### SYS-P2-CONCURRENCY-SCOPE — OPEN_OPTIMIZATION

M3/M4 currently share a conservative global `PROTOCOLLO` token. This can cause unnecessary stale/refresh behavior, but is not a demonstrated lost-update defect and is not a reason for preventive churn before the native gate.

### RC-PACKAGE — PENDING

Do not create/freeze the final RC package until the native gate is complete. The final RC artifact must be consolidated, re-extracted, hash-verified and re-tested as one installable package.

## Mandatory native gates before RC freeze

- [ ] `Debug > Compila VBAProject`
- [ ] real install/upgrade on a clean Beta copy
- [ ] real M3↔M4 stale/reload/apply scenarios
- [ ] Workbook events / `cMinimalSaveEvents`
- [ ] Save → close → reopen
- [ ] FSC identical/new/blocking scenarios
- [ ] real PDF `ExportAsFixedFormat`
- [ ] two-workbook `Shape.OnAction` dispatch
- [ ] representative dataset performance/memory
- [ ] real reentrancy attempt during `AGGIORNA_TUTTO` across `DoEvents`
- [ ] STEP8B read-only audit on the real operational `.xlsm`

## Promotion decision rule

The `beta -> main` PR must remain **Draft / DO NOT MERGE** until all mandatory native gates are recorded as PASS or a specific item is deliberately deferred with an explicit, technically justified non-RC-blocking rationale.

A new STEP10B remediation is **not** warranted merely for preventive code churn. Create a further remediation checkpoint only for a reproducible defect discovered by the native Excel/VBE gate or the real-workbook audit.

## Repository scope and data safety

The operational `.xlsm` containing real user data must never be committed. Public repository documentation must not imply that cloning GitHub yields the complete operational workbook unless a sanitized complete source/package is actually published.

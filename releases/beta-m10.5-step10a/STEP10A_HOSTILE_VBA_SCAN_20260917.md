# WebArch — STEP 10A Hostile VBA Scan

Date: 17 September 2026
Branch: `beta`
Checkpoint: **M10.5 / M11.5 — STEP 10A**

## Verdict

**HOSTILE OFFLINE VBA SCAN: PASS — no new confirmed P0/P1/P2 defects found.**

Release state remains **BETA OFFLINE VERIFIED / RC BLOCKED** because the Microsoft Excel/VBE native gate and the real operational workbook gate are still pending.

## Scope

The hostile scan resumed from the previous stop point and reviewed the cumulative STEP10A source state, preserving the accepted Alpha baseline `EASYONE MINIMAL 1.0.3 + HOME 1.1` unchanged.

Reference source state:

- 91 VBA modules;
- STEP9 independent re-audit: 320/320 PASS;
- STEP10A full static audit: 643/643 PASS;
- STEP10A suite: 19/19 PASS;
- STEP10A specific gate: 24/24 PASS;
- STEP10A modified modules: 7/91, limited to versioning/naming normalization.

## Hostile checks completed

The final pass concentrated on VBA failure modes that can escape ordinary syntax/static checks:

- `And` / `Or` non-short-circuit evaluation;
- `IIf` eager evaluation;
- unsafe `Nothing` dereferences;
- broad or misleading `On Error Resume Next` scopes;
- error preservation before cleanup;
- EASYONE acquisition/release symmetry;
- Excel application-state capture/restore symmetry;
- reentrancy windows and `DoEvents` exposure;
- dynamic `Application.Run` / call-contract risks;
- `ActiveWorkbook`, `ActiveSheet`, `Selection` style context ambiguity;
- `ListObject` / `DataBodyRange` assumptions;
- Save / SaveAs / final-save paths;
- discard / last-good / last-attempt fail-closed behavior;
- FSC finalization semantics;
- PDF freshness/provenance checks;
- M10 wrapper cleanup and release paths;
- regression risk introduced by STEP10A.

## Findings

### New confirmed defects

- P0: **0**
- P1: **0**
- P2: **0**

### Regression findings

- STEP10A regressions confirmed: **0**

### Previously known `Nothing Or property` defects

The historical M4 pattern such as:

`If lo Is Nothing Or lo.DataBodyRange Is Nothing Then`

was a real VBA bug because `Or` does not short-circuit. The current remediation line already replaces it with sequential guards and validates the expected `ListObject`. No reintroduction was found.

### `IIf` residual risk in frozen Alpha

A previously known `IIf` use in frozen Alpha `modBacklog` remains classified **LOW / LATENT_NOT_REPRODUCED**. It is a language-level fragility because VBA evaluates both branches, but no failure has been reproduced with the admitted input path. This scan does **not** justify modifying the frozen Alpha.

### False-positive distinction

Conditions such as:

`obj1 Is Nothing Or obj2 Is Nothing`

are not equivalent to `obj Is Nothing Or obj.Property ...`: evaluating `Is Nothing` on each object does not itself dereference an object property. The hostile scan treated only property/member access after a potentially `Nothing` object as a defect candidate.

## Open items that are not new VBA bugs

1. `SYS-P2-CONCURRENCY-SCOPE` — **OPEN_OPTIMIZATION**. M3/M4 share a conservative global `PROTOCOLLO` token. This may cause unnecessary stale/refresh events but is not a lost-update defect.
2. `SURFACE-P1-01` — **PENDING_REAL_WORKBOOK_GATE**. Macro-surface reduction remains forbidden until the real `.xlsm` is audited for persisted `Shape.OnAction` bindings.
3. `NATIVE-GATE` — **PENDING**. Static/offline evidence cannot replace Excel/VBE execution.

## Native gates still required before RC freeze

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
- [ ] STEP8B read-only binding audit on the real operational `.xlsm`

## Decision

The hostile offline VBA scan is **CLOSED** for STEP10A.

No STEP10B remediation should be created merely for preventive code churn. A STEP10B is justified only if a reproducible defect emerges from the native Excel/VBE gate or the real-workbook binding audit.

The PR `beta -> main` must remain **Draft / DO NOT MERGE** until the native gate is closed.
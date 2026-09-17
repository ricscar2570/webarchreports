# WebArch — RC Promotion Gate

**Branch:** `beta`  
**Current checkpoint:** `M10.5 / M11.5 — STEP 6B`  
**Status:** `OFFLINE VERIFIED / NOT RC FROZEN`  
**Stable accepted baseline:** `EASYONE MINIMAL 1.0.3 + HOME 1.1` on `main` / `alpha-frozen-1.0.3-home1.1`

This document is the merge gate for promoting the active Beta line into `main`.

## Branch roles

- `alpha-frozen-1.0.3-home1.1`: immutable historical freeze of the accepted Alpha record.
- `main`: stable/public baseline registry. It must not silently become a Beta development branch.
- `beta`: active remediation and validation line. It may advance while RC gates remain open.

No force-push should be used to rewrite the published milestone history.

## Current verified Beta checkpoint

The current Beta registry identifies:

- milestone: `M10.5-M11.5-step6b-offline-verified`;
- package: `WebArch_M10_5_STEP6B_P1_05_REENTRANCY.zip`;
- SHA-256: `021a55369e87c56419bbec22f29658052552e990704ba3bdad4efedb7b157e0f`;
- cumulative suites: `17/17 PASS`;
- release verifier: `92/92 PASS`;
- static / model / integration evidence: PASS as recorded in `releases/beta-m10.5-step6b/QA_SUMMARY.md`.

This evidence is offline/static/model-based. It is not a substitute for Microsoft Excel/VBE execution.

## Hard blockers before RC freeze

All items below must be explicitly recorded as PASS (or, where applicable, deliberately DEFERRED with rationale) before this Beta can be promoted to `main`.

### Native Excel/VBE gates

- [ ] `NATIVE-01` — `Debug > Compila VBAProject` succeeds on the candidate package.
- [ ] `NATIVE-02` — real install/upgrade path on a clean Beta copy succeeds.
- [ ] `NATIVE-03` — M3↔M4 stale/reload/apply scenarios succeed without lost update.
- [ ] `NATIVE-04` — Workbook events / `cMinimalSaveEvents` behave correctly.
- [ ] `NATIVE-05` — Save → close Excel → reopen preserves required state.
- [ ] `NATIVE-06` — FSC identical/new/blocking scenarios behave as designed.
- [ ] `NATIVE-07` — PDF `ExportAsFixedFormat` succeeds and output is valid.
- [ ] `NATIVE-08` — two open WebArch workbooks dispatch every button to the owning workbook.
- [ ] `NATIVE-09` — performance/memory are acceptable on a representative real-sized dataset.
- [ ] Reentrancy test during `AGGIORNA_TUTTO` crosses real `DoEvents`; protected M10/M4 commands are refused while EASYONE is busy and work normally after EASYONE returns `FREE`.

### Audit reconciliation gates

The 16 September total-application audit originally reported additional P1/P2 items beyond the P0/P1 findings now documented as closed offline. Before RC freeze, each must have an explicit disposition in the release record; absence from a later README is not by itself proof of closure.

- [ ] Macro surface (`APP-P1-06` / `SURFACE-P1-01`) classified: supported UI vs admin/test/recovery/legacy.
- [ ] PDF filename/version provenance (`APP-P2-01`) verified or explicitly closed.
- [ ] PDF freshness check after SaveAs and immediately before export (`APP-P2-02`) verified or explicitly closed.
- [ ] M3/M4 concurrency-scope optimization (`APP-P2-03` / `SYS-P2-01`) either implemented or consciously deferred as a non-RC blocker with rationale.
- [ ] Hardening/version naming consistency (`APP-P2-04` / `DOC-P2-01`) reconciled.
- [ ] Procedure-scoped concurrency verifier finding (`APP-P1-02` / `QA-P1-01`) explicitly marked CLOSED with the exact verifier evidence that replaced the earlier file-global check.

## Repository/package gates

- [ ] Candidate source identity is frozen and recorded.
- [ ] Candidate package SHA-256 is frozen and recorded.
- [ ] Modified-file manifest is complete.
- [ ] Cold-overlay/reconstruction result matches the audited candidate byte-for-byte for all VBA modules in scope.
- [ ] No operational workbook containing real user data is committed.
- [ ] Public-source scope is documented accurately; do not claim that cloning this repository yields the full operational product unless the complete sanitized source/package is actually published.

## Promotion procedure

Only after all hard blockers are resolved:

1. create the RC freeze record under `releases/`;
2. update `VERSION.json` from `release_candidate_frozen: false` to the frozen RC identity;
3. record all native-test evidence and residual-finding dispositions;
4. verify package and manifests one final time;
5. review the `beta -> main` pull request;
6. merge without rewriting history;
7. tag the promoted commit/release with the frozen RC version.

Until then, the `beta -> main` pull request must remain a **Draft / DO NOT MERGE** promotion candidate.

## Non-regression rule

The accepted Alpha remains untouched. Beta remediation must preserve all previously closed fixes listed in the correction ledger and must not reintroduce the old transactional overlay, BUSY nesting, last-good overwrite, weak fingerprint/token logic, unsafe FSC finalization, or unprotected standalone wrappers.

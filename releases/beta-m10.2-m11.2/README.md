# WebArch Beta M10.2 / M11.2 — Full Remediation

Status: **gate candidate; RC not frozen**.

M10.2 supersedes M10.1 after the 16 Sep 2026 full VBA/flow re-audit. The remediation removes the generic dynamic M3–M9 invoker and uses typed API calls, adds the complete M4 pending-edit gate, a controlled EASYONE/MINIMAL safety envelope for standalone realignment and FSC finalization, day/KPI staleness, FP3 value+formula fingerprints, migration blocking from older M10 versions, and hardened report/KPI handling.

M4 source hardening remains mandatory before RC freeze: use the supplied M4H2 source patcher on the exact exported M4 module, then compile and rerun the M4 gate. The additive M4H2 guard already protects coordinated M10.2 paths.

FSC behavior is deliberately conservative: M10.2 declares/auto-saves a new FSC publication only when `FSC_DB_CORRENTI` changes. If the current snapshot is unchanged, diagnostics/status/history may still have changed, but M10.2 makes no success claim and performs no automatic Save.

Offline verification:

- Static QA: 212/212 PASS
- Compile-oriented QA: 269/269 PASS
- Typed call contracts: 14 calls, 0 errors
- State/failure-path models: 51/51 PASS
- Negative mutations: 14/14 PASS
- M4H2 patcher: 11/11 PASS
- Final ZIP manifests and ZIP integrity: PASS

Microsoft VBE compile, Excel runtime, real import, Save/reopen, PDF and performance gates remain **NOT EXECUTED** and are required before RC1.

Exact binary/source delivery is archived in the WebArch Project Library; this repository path records status, gates and checksums.
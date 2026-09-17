# WebArch — System integration status M10.4

| Shared state | Writer / consumer | Protection | Offline result |
|---|---|---|---|
| M2 PROTOCOLLO | M3 | applied + PENDING baseline before Apply | PASS |
| M2 PROTOCOLLO | M4 | applied + PENDING baseline before Apply | PASS |
| M3 ↔ M4 | two editors | both reject Apply after M2 changes | PASS |
| M2 PROTOCOLLO → TASK | M5 sync | derived consumer; does not rewrite PROTOCOLLO | PASS |
| M5 TASK | OP_SCADENZE | technical columns validated before commit | PASS |
| M5 TASK | reminder sync | manual fields preserved on existing tasks | PASS |
| M6 DECRETO | M7 | loaded / expected revision | PASS |
| manual DECRETO store | M7 | loaded manual revision | PASS |
| M6 INCIDENTE | M8 | expected rectification revision | PASS |
| manual INCIDENTE store | M8 | loaded manual revision | PASS |
| M2 + M6 FASCICOLO | M9 | loaded user + rectification revision | PASS |
| Alpha DB | M3/M4/M7/M8 | source columns validated before commit | PASS |
| B1 current snapshot | M9 | source identity + revision guards | PASS |
| OP views / KPI | M10.4 | deterministic order + mandatory refresh | PASS |
| partial writes M5/M7/M8 | Save | safety envelope + discard | PASS |
| legacy macros | M10 state | source/state token | PASS |
| day change | M10 state | DAY in token | PASS |
| report KPI | M10 state | KPI signature included | PASS |

**Invariant:** no view that writes shared state may apply values if the state on which it was built changed after its last refresh.

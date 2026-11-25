# Engineering Specification Review

- Project: Atlas Memory Interface
- Revision: B
- Interface: Synthetic DDR-style interface
- Owner: Engineering documentation
- Parameters: 6 \| Pass: 5 \| Fail: 1

![Workflow](workflow.png)

| Parameter | Value | Unit | Minimum | Maximum | Status |
| --- | --- | --- | --- | --- | --- |
| Supply voltage | 1.10 | V | 1.05 | 1.15 | PASS |
| Clock frequency | 800 | MHz | 750 | 850 | PASS |
| Bus width | 64 | bit | 64 | 64 | PASS |
| Operating temperature | 90 | C | 0 | 85 | FAIL |
| Current draw | 20 | mA | 10 | 30 | PASS |
| Setup time | 0.25 | ns | 0.20 | 0.50 | PASS |

Synthetic demonstration - engineering review required

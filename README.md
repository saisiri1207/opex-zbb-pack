# OpEx ZBB pack

Zero-based OpEx pack for a fictional CPG company (**Northline Consumer Products**): cost-center asks, decision packages (Must / Should / Nice), fund Y/N toggles, and a Prior → Ask → ZBB cuts → Target savings bridge.

**File to open:** [`Northline_OpEx_ZBB_Pack.xlsx`](Northline_OpEx_ZBB_Pack.xlsx)

`build.py` only regenerates that workbook. The file a hiring manager should open is the `.xlsx`.

## What you will see

- Cost centers with prior-year actuals, FY26 ask, and FTE drivers
- Decision packages rebuilt by tier (**Must-Have / Should-Have / Nice-to-Have**)
- Fund Y/N toggles that roll into a recommended budget
- Savings bridge: Prior → Ask uplift → tier cuts → Recommended vs CFO target
- Dashboard: ask vs target, tier mix, gap-to-target

Yellow cells with blue font are inputs. Black font is formulas. Amounts in $000s.

## How to use (≈8 minutes)

1. Open `01_Assumptions`. Set the FY26 OpEx target and skim ZBB policy.
2. Review prior / ask / FTE on `02_Cost_Centers`.
3. Fund or defer packages on `03_Decision_Packages` (yellow Fund Y/N).
4. Walk the OpEx bridge on `04_Savings_Bridge` — check vs target.
5. Close on `05_Dashboard`: tiles, comparison chart, funded mix by tier.

## Screen-share test

On `03_Decision_Packages`, flip a funded Should-Have from Y → N. Recommended OpEx and gap-to-target on the Dashboard should move; the Should-Have cut step on the bridge deepens. Flip it back — numbers reverse.

## Tabs

| Tab | Role |
| --- | --- |
| `00_Cover` | Purpose, legend, fictional disclaimer |
| `01_Assumptions` | Target, ZBB policy, tier legend, roll-up checks |
| `02_Cost_Centers` | Prior, ask, Δ, FTE drivers by cost center |
| `03_Decision_Packages` | Package register, Fund Y/N, tier rollup |
| `04_Savings_Bridge` | Prior → Ask → cuts → Recommended vs target |
| `05_Dashboard` | Ask vs target tiles, charts, tier mix |
| `06_Data_Dictionary` | Field definitions |

Excel formulas only. No VBA, no live ERP feed, no employer data. This is the *OpEx / ZBB / decision-package* slice — not a full three-statement model.

Fictional company and sample data for portfolio demonstration only.

[Profile](https://github.com/saisiri1207) · [Portfolio](https://saisiri1207.github.io) · [LinkedIn](https://www.linkedin.com/in/saisiri1207) · [bandarusaisiri1207@gmail.com](mailto:bandarusaisiri1207@gmail.com)

Sai Siri Bandaru — Financial Analyst | FP&A

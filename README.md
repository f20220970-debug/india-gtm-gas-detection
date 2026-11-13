# India Go-to-Market Model: Portable Gas Detectors for Factories

A 3-year go-to-market plan and financial model for launching a portable 4-gas detector (oxygen, flammable gas, CO, H₂S) in India. It answers: **who to sell to first, at what price, through which channel, and when it breaks even.** It plays to my chemistry background: these devices decide whether a worker can safely enter a tank or reactor.

Built by Pavitra Seth (BITS Pilani, B.E. Mechanical + M.Sc. Chemistry) as preparation for Honeywell PS-II. It uses public data only and is not affiliated with Honeywell.

**Files**
- `india_gtm_model.xlsx`: tabs README, Inputs, Segments, Pricing, Channel, Plan (quarterly, 3 years), UnitEconomics, Sensitivity, Launch, Sources. It has 348 live formulas and no errors.
- `build.py`: rebuilds the workbook (`python build.py`).

## Who to sell to first
| Segment | Factories in operation (ASI 2023-24) | Target factories | Annual detector market (units) |
|---|---|---|---|
| **Chemicals** | 12,148 | 9,718 | 6,414 |
| Pharmaceuticals | 4,923 | 3,446 | 1,882 |
| Refined petroleum | 1,538 | 1,538 | 7,167 |
| Basic metals | 10,834 | 3,250 | 2,145 |
| Food: ammonia cold storage | 35,081 | 1,754 | 460 |
| **Total** | | **19,707** | **18,067 units ≈ ₹58 Cr a year** |

**Start with chemicals.** The need is high, the factories are clustered (Vapi–Ankleshwar–Dahej), the Hazardous Chemical Rules apply, and many units are MSMEs that global brands serve poorly. Refineries need more detectors per site but already buy global brands directly. As a cross-check, the 19,707 target factories are 43% of India's 45,374 registered hazardous-process factories (DGFASLI, 2023).

## Price
Market price points on IndiaMART range from ₹12,500 (unbranded) to ₹70,000 (Industrial Scientific), with a median of ₹28,000. Chosen tiers:

| Tier | Price | Share of sales | Position |
|---|---|---|---|
| Value | ₹18,000 | 30% | Certified and serviced, against unbranded units at ₹12.5–13.5k |
| Standard | ₹32,000 | 60% | Below the Honeywell BW Flex (₹39k) and Dräger (₹46k) |
| Connected | ₹52,000 | 10% | Audit logs for pharma exporters |

Recurring revenue comes from annual calibration (₹4,000 per device) and sensor replacement (₹8,000 every 2 years). The pitch is risk, not specs: one avoided 10-day shutdown is worth about **25×** a 4-detector order.

## Results
| | Year 1 | Year 2 | Year 3 |
|---|---|---|---|
| Revenue | ₹1.13 Cr | ₹3.77 Cr | ₹7.35 Cr |
| Gross margin | 50% | 48% | 48% |
| Operating profit | −₹0.80 Cr | ₹0.15 Cr | ₹1.12 Cr |
| New customers | 90 | 270 | 450 |

- First profitable quarter: **Year 2, Q2**. Cumulative payback: **Year 3, Q3**. Peak cash need: **₹83 lakh**.
- CAC is **₹40,000**, 34% of the first order. LTV/CAC is **3.6×**, and the first order's gross profit already covers CAC.

## Three takeaways
1. **Win rate matters more than price.** Moving demo-to-order from 30% to 25% cuts Year-3 profit by 40%. Moving price from ₹32k to ₹28k cuts it by 20%. So invest in free trials and demos before discounting.
2. **The market is focused but small** (about ₹58 Cr a year at standard price). By Year 3 the plan needs a 13% unit share, above my 10% comfort level. Growth beyond that needs fixed detectors or adjacent segments (pharma exporters, cold chains).
3. **Recurring service is the moat against cheap imports.** Calibration and sensors grow to a fifth of revenue by Year 3 and make the customer relationship sticky. Unbranded sellers don't offer them.

## Launch plan
- **Pilot (days 0–90):** one chemical cluster, 2 sales engineers, 1 distributor and 20 free trials. Gate to move on: 30 paying customers and at least 40% of trials converting.
- **Year 2:** add the Hyderabad and Baddi pharma clusters, two more distributors, and a GeM listing.
- **Year 3:** five clusters, plus the Connected tier for exporters.
- **Risks and mitigations** are on the Launch tab: price war, certification delays, distributor conflict, customers skipping calibration, and sensor supply.

## Limitations
- Factory counts are ASI sample estimates.
- Shares, detectors per factory, funnel rates and costs are assumptions (yellow cells) to test in 20 customer interviews.
- Prices are list prices from one marketplace on one date.

## Sources
MoSPI Annual Survey of Industries 2023-24, DGFASLI (via Dataful), IndiaMART listings, CPCB (Hazardous Chemical Rules), PESO and GeM. Full links are on the Sources tab.

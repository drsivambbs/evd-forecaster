# EVD Forecaster — What We Built

## What the tool is

We built a web app, the **EVD Forecaster** (EVD = Ebola Virus Disease). It turns raw outbreak case numbers into clear answers to three questions:

- *Is the outbreak growing or shrinking right now?*
- *How many more cases should we expect in the coming months?*
- *When is it likely to end, and how much does that depend on our control efforts?*

It runs as a simple **5-step workflow** in the browser. Every input and its scientific source appears on screen and can be edited. The core engine is the **Cori et al. (2013)** method (EpiEstim): it estimates how many people one infected person infects each day (above 1 means spreading, below 1 means fading). This is the same method the WHO Ebola Response Team used in 2014.

---

## Step 1 — Building the daily case series

WHO bulletins report cumulative totals on irregular dates, but the model needs **new cases per day**. Step 1 converts cumulative totals (or daily counts) into a clean daily series across three WHO categories: **confirmed**, **suspected**, and **deaths**. It fills the gaps between reporting dates by linear interpolation and reconstructs the earliest days by back-extrapolation, so the totals reconcile and no cases are lost. Each row records its source for traceability.

## Step 2 — Estimating the reproduction number (R_t)

Using the daily series, the tool applies the **Cori et al. (2013)** method to estimate R_t for each day (above 1 the outbreak grows, below 1 it fades). It reports R_t as a **95% credible interval** (Bayesian posterior, not a frequentist confidence interval), so the uncertainty is shown directly on the chart. A **sensitivity check** re-runs the calculation with a second serial interval (the typical gap between successive cases); if both lines agree, the conclusion is robust. R_t is plotted over time with a reference line at 1.

## Step 3 (optional) — The natural course (SEIHFR compartmental model)

The renewal-equation forecast in Step 4 requires you to assume how R_t will fall, so it cannot show what the outbreak would do on its own. Step 3 answers that. It is a six-compartment model that moves people through **S**usceptible → **E**xposed → **I**nfectious → **H**ospitalised → **F**uneral → **R**ecovered states (SEIHFR), based on **Legrand et al. (2007)**.

It models three transmission routes separately: **community (β_I)**, **hospital (β_H)**, and **funeral (β_F)**. The funeral route matters because Ebola spreads strongly from bodies during burial rituals (**Wamala et al., 2010** — adjusted odds ratio 3.83). The tool **fits the community rate β_I** to your observed confirmed cases and **anchors β_H and β_F** on the Wamala odds ratios, so the three rates can never collapse to zero with sparse data. The natural-history durations (incubation, onset-to-hospitalisation, onset-to-death, hospitalisation-to-death, onset-to-recovery, death-to-burial) set the transition rates between states; defaults are Bundibugyo values from Wamala 2010 with Legrand 2007 used as a Zaire proxy where Bundibugyo-specific values are not published.

**Why we need it:** unlike the renewal forecast, this model tracks the susceptible pool, so the outbreak naturally rises, peaks, and falls as susceptibles run out. This gives a true **hands-off baseline** — the natural course — with no R_t scenario imposed. You then compare it against intervention scenarios by reducing the relevant transmission rate:
- **Natural** (no action) — all rates at baseline
- **Funeral control** (safe burials) — β_F cut to 20%
- **Combined** (safe burials + PPE + isolation) — β_H cut to 40%, β_F cut to 20%

The difference between the natural curve and the intervention curves measures how many cases and deaths each control measure prevents.

## Step 4 — Forecasting future cases

Starting from the R_t selected in Step 2 (or a user-defined value), the **renewal-equation forecast** (**Nouvellet et al., 2018**) projects cases and deaths forward over a selected horizon (default 365 days). The model evaluates three response scenarios in parallel:
- **S1 — Delayed response**
- **S2 — Moderate response**
- **S3 — Strong combined response**

Each scenario is defined by a **target R_t value** and the **number of days required to reach it**. The delayed scenario reaches the target more slowly, illustrating the operational cost of a slower response. Death projections are derived using the **case fatality ratio (CFR)** and an **onset-to-death delay** (defaults from Wamala 2010 Bundibugyo: CFR = 0.34, delay = 10 days). For each scenario, results are presented with a **90% posterior predictive interval** obtained by sampling 200 R_t starting values from the Cori Gamma posterior of Step 2.

## Step 5 — Predicting the end of the outbreak

After the last reported case, how long before we can declare the outbreak over? The danger is the cases we can't see — **under-reporting** (typically only 10–30% detected) and **superspreading** (a rare few infect many). The tool runs a **stochastic offspring-tree simulation** (Nishiura framework with **Lloyd-Smith et al., 2005** superspreading dispersion k = 0.18 for EBOV), replaying the descendant chain a thousand times to build a day-by-day probability that the outbreak is truly over. It reports the date when confidence passes **95%** and compares this with the **WHO 42-day rule** and the stricter **Djaafara et al. (2021) 63-day preliminary + 90-day final** declaration framework.

---

## Technical notes

Built in **Python 3.10+** with **Streamlit** (pandas, NumPy, SciPy, Plotly, openpyxl, kaleido).

- **Source code:** [github.com/drsivambbs/evd-forecaster](https://github.com/drsivambbs/evd-forecaster)
- **Live app:** [evd-forecaster-version2.streamlit.app](https://evd-forecaster-version2.streamlit.app)

Runs in any modern browser. Every run can be exported as a **formatted Excel (.xlsx) report** containing a cover sheet, all inputs with DOI links, the four charts (embedded as PNG via kaleido), per-step results, and an auto-generated interpretation narrative.

## Current limitations

- It assumes the **cumulative case data in WHO Disease Outbreak News (DONs) is sufficient**; daily line-list data would give finer accuracy.
- **R_t trajectories in Step 4** are user-defined scenarios — they reflect response-speed assumptions, not predictions of what will happen.
- **Deaths are derived deterministically** from confirmed cases with a fixed onset-to-death lag; the true onset-to-death distribution is right-skewed and the death wave shape is approximate.
- **Parameter defaults are tuned for Bundibugyo Ebola**. For Zaire EBOV, Sudan, Mpox, or other pathogens, the CFR, serial interval, dispersion k, and stage durations must be replaced.
- **Final bug clearance and statistical verification are still pending** and should be completed before operational deployment.

---

## References (APA, short form)

1. **Cori, A., Ferguson, N. M., Fraser, C., & Cauchemez, S.** (2013). A new framework and software to estimate time-varying reproduction numbers during epidemics. *American Journal of Epidemiology*, 178(9), 1505–1512. [doi.org/10.1093/aje/kwt133](https://doi.org/10.1093/aje/kwt133)
2. **Nouvellet, P., Cori, A., Garske, T., et al.** (2018). A simple approach to measure transmissibility and forecast incidence. *Epidemics*, 22, 3–12. [doi.org/10.1016/j.epidem.2017.02.012](https://doi.org/10.1016/j.epidem.2017.02.012)
3. **WHO Ebola Response Team.** (2014). Ebola virus disease in West Africa — the first 9 months of the epidemic and forward projections. *New England Journal of Medicine*, 371(16), 1481–1495. [doi.org/10.1056/NEJMoa1411100](https://doi.org/10.1056/NEJMoa1411100)
4. **Lloyd-Smith, J. O., Schreiber, S. J., Kopp, P. E., & Getz, W. M.** (2005). Superspreading and the effect of individual variation on disease emergence. *Nature*, 438, 355–359. [doi.org/10.1038/nature04153](https://doi.org/10.1038/nature04153)
5. **Wamala, J. F., Lukwago, L., Malimbo, M., et al.** (2010). Ebola hemorrhagic fever associated with novel virus strain, Uganda, 2007–2008. *Emerging Infectious Diseases*, 16(7), 1087–1092. [doi.org/10.3201/eid1607.090536](https://doi.org/10.3201/eid1607.090536)
6. **Legrand, J., Grais, R. F., Boelle, P. Y., Valleron, A. J., & Flahault, A.** (2007). Understanding the dynamics of Ebola epidemics. *Epidemiology and Infection*, 135(4), 610–621. [doi.org/10.1017/S0950268806007217](https://doi.org/10.1017/S0950268806007217)
7. **Djaafara, B. A., Imai, N., Hamblion, E., et al.** (2021). A quantitative framework for defining the end of an infectious disease outbreak: application to Ebola virus disease. *American Journal of Epidemiology*, 190(4), 642–651. [doi.org/10.1093/aje/kwaa249](https://doi.org/10.1093/aje/kwaa249)

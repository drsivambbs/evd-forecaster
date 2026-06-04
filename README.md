# EVD Forecaster

**An interactive Ebola-outbreak modelling app — written so a first-year MPH student can use it.**

This is a simple web app that helps public-health workers and students answer five basic questions about an Ebola outbreak:

1. *What does the outbreak look like day by day?* — Step 1
2. *Is it growing or shrinking right now?* — Step 2
3. *What would happen if no one intervened?* — Step 3 (optional)
4. *How big could it get under different response speeds?* — Step 4
5. *When can we say the outbreak is over?* — Step 5

You enter a few numbers from official reports (WHO Disease Outbreak News, etc.), the app does the maths, and you get charts and tables you can put straight into a report. Every input has a clickable citation (DOI) so anyone reviewing your work can check where each number came from.

---

## Why each step exists — plain language

| Step | What it does | The question it answers |
|---|---|---|
| **1. Daily incidence builder** | Turns cumulative totals from WHO reports (e.g., "134 confirmed cases by May 29") into a day-by-day series. | *How many new cases per day?* |
| **2. R_t estimation** | Calculates the **reproduction number** R_t — the average number of new infections per current case. | *Is each sick person passing it to more than 1 other person (R_t > 1, growing) or fewer (R_t < 1, declining)?* |
| **3. SEIHFR model (optional)** | Runs a classic 6-compartment model (Susceptible → Exposed → Infectious → Hospitalised → Funeral → Recovered) for the *natural* outbreak with no response. | *What would the outbreak look like if nobody did anything? How big could it possibly get?* |
| **4. Renewal-equation forecast** | Projects the outbreak forward 365 days under **three response scenarios**: Delayed, Moderate, Strong. | *How does the outbreak change depending on how fast the response is?* |
| **5. End-of-outbreak (EOO) predictor** | Calculates the probability that the outbreak has truly ended, T days after the last case. | *When can WHO declare the outbreak over?* |

---

## How to use it

### Option A — Use the live web app
Open the deployed app (Streamlit Community Cloud link). Nothing to install.

### Option B — Run it on your own computer
```bash
git clone https://github.com/drsivambbs/evd-forecaster.git
cd evd-forecaster
pip install -r requirements.txt
python -m streamlit run app.py
```
Then open the URL shown (default `http://localhost:8501`).

### A typical workflow
1. **Step 1**: Enter cumulative case numbers from 3–4 WHO DON reports. The app auto-converts them into a daily series.
2. **Step 2**: Click **Run R_t**. You'll see a chart with a red line (R_t mean) and a shaded credible interval. If the line is above the dotted "R_t = 1" line, the outbreak is growing.
3. **Step 2 → Pick an R_t** for the forecast (the **Latest** value is recommended).
4. **Step 3 (optional)**: Run SEIHFR if you want to see the "what if no response" curve as a worst-case ceiling.
5. **Step 4**: Click **Run forecast**. You get three coloured curves — Delayed, Moderate, Strong response. The shaded bands are 90% prediction intervals.
6. **Step 5**: Click **Run EOO**. You'll see the probability of outbreak-over rising over time, with markers at the WHO 42-day and Djaafara 63-day rules.
7. **Top-right corner**: Click **Generate Excel report** to download a polished `.xlsx` with cover page, inputs, charts, results, and interpretation.

---

## Key indicators explained (the four-line versions)

- **R_t (reproduction number)** — on average, how many new people each sick person infects right now. > 1 = growing; < 1 = shrinking.
- **Serial interval (SI)** — average days between one person showing symptoms and the next person they infected showing symptoms. For Ebola: ~15 days.
- **CFR (case fatality ratio)** — out of every 100 confirmed cases, how many die. For Bundibugyo Ebola: about 34.
- **Attack rate** — fraction of the population that gets infected over the whole outbreak. Higher R₀ = higher attack rate.

For more, see the built-in **Input glossary** button in the app (top-right).

---

## What's under the hood (the technical short version)

| Step | Method | Source |
|---|---|---|
| 2. R_t | Cori 2013 Bayesian sliding-window with Gamma posterior | Cori 2013 |
| 3. SEIHFR | 6-compartment ODE with β_F/β_I and β_H/β_I ratios anchored on Wamala 2010 funeral OR | Legrand 2007 + Wamala 2010 |
| 4. Forecast | Renewal equation I_t = R_t · Σ w_s · I_{t-s}; 200 R_t samples for uncertainty bands | Nouvellet 2018 |
| 5. EOO | Stochastic branching process with Negative-Binomial(R, k=0.18) offspring | Nishiura framework + Lloyd-Smith 2005 |

Each step shows the actual equations (LaTeX-rendered) in a collapsible **📐 Method & equations** panel at the bottom of the page.

---

## Important limitations (please read before quoting results)

- **Linear interpolation between cumulative snapshots** invents daily counts. If you have real daily surveillance data, use that instead.
- **The "Natural" SEIHFR scenario** assumes a fully susceptible population with no intervention. Real outbreaks always trigger *some* response, so its peak numbers are an upper-bound thought experiment, not a prediction.
- **Forecast scenarios (Delayed/Moderate/Strong)** are operational assumptions about R_t decline — you choose the speed. They are not forecasts of what *will* happen.
- **CFR is applied with a fixed 10-day lag** to projected cases. Real onset-to-death distributions are right-skewed, so the death wave shape is approximate.
- **Parameter defaults are tuned for Bundibugyo Ebola** (Wamala 2010). For Zaire EBOV, Sudan, Mpox, or other pathogens you must replace CFR, SI, k, and stage durations.
- **R_t is right-aligned** to the window end (Cori convention). It reflects transmission centred about τ/2 days earlier.

---

## Key references (clickable DOIs)

- **Cori et al. 2013** — Sliding-window Bayesian R_t estimator. *Am J Epidemiol* 178:1505. [10.1093/aje/kwt133](https://doi.org/10.1093/aje/kwt133)
- **Nouvellet et al. 2018** — Renewal-equation short-term forecasting. *Epidemics* 22:3. [10.1016/j.epidem.2017.02.012](https://doi.org/10.1016/j.epidem.2017.02.012)
- **WHO Ebola Response Team 2014** — Serial-interval parameters from West Africa. *N Engl J Med* 371:1481. [10.1056/NEJMoa1411100](https://doi.org/10.1056/NEJMoa1411100)
- **Lloyd-Smith et al. 2005** — Superspreading & offspring-distribution dispersion k. *Nature* 438:355. [10.1038/nature04153](https://doi.org/10.1038/nature04153)
- **Wamala et al. 2010** — Bundibugyo Ebola natural history + funeral / hospital ORs. *Emerg Infect Dis* 16:1087. [10.3201/eid1607.090536](https://doi.org/10.3201/eid1607.090536)
- **Legrand et al. 2007** — SEIHFR compartmental framework + Zaire-proxy stage durations. *Epidemiol Infect* 135:610. [10.1017/S0950268806007217](https://doi.org/10.1017/S0950268806007217)
- **Djaafara et al. 2021** — Reporting-rate-adjusted EOO declaration (63 + 90 days). *Am J Epidemiol* 190:642. [10.1093/aje/kwaa249](https://doi.org/10.1093/aje/kwaa249)

---

## Files in this repo

| File | What it is |
|---|---|
| `app.py` | The full Streamlit application (single file) |
| `requirements.txt` | Python packages needed |
| `.gitignore` | Files git should ignore (Python cache, virtualenvs, secrets) |
| `README.md` | This file |

---

## Contributing & feedback

Issues and pull requests welcome at [github.com/drsivambbs/evd-forecaster](https://github.com/drsivambbs/evd-forecaster).
If you're an MPH student using this for coursework and something is unclear, that's a bug — please open an issue.

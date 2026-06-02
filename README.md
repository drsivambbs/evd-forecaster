# EVD Forecaster

A four-step interactive Streamlit app for Ebola Virus Disease (EVD) outbreak modelling:

1. **Daily incidence builder** — convert cumulative DON snapshots or sparse incidence data into a clean daily case series.
2. **R<sub>t</sub> estimation** — instantaneous reproduction number using the Cori et al. (2013) Bayesian sliding-window method.
3. **Renewal-equation forecast** — Nouvellet et al. (2018) projection under three response scenarios (Delayed / Moderate / Strong combined), with 90% posterior predictive bands.
4. **End-of-outbreak predictor** — Nishiura / Lloyd-Smith descendant-tree simulation with WHO 42-day and Djaafara 63-day declaration thresholds.

Every input is sourced and editable. A built-in glossary explains each input twice — once for modellers, once in plain English for MPH students.

## Run locally

```bash
pip install -r requirements.txt
python -m streamlit run app.py
```

Open the URL shown (default `http://localhost:8501`).

## Key references

- Cori A et al. (2013) *Am J Epidemiol* 178:1505. [10.1093/aje/kwt133](https://doi.org/10.1093/aje/kwt133)
- Nouvellet P et al. (2018) *Epidemics* 22:3. [10.1016/j.epidem.2017.02.012](https://doi.org/10.1016/j.epidem.2017.02.012)
- WHO Ebola Response Team (2014) *N Engl J Med* 371:1481. [10.1056/NEJMoa1411100](https://doi.org/10.1056/NEJMoa1411100)
- Lloyd-Smith JO et al. (2005) *Nature* 438:355. [10.1038/nature04153](https://doi.org/10.1038/nature04153)
- Wamala JF et al. (2010) *Emerg Infect Dis* 16:1087. [10.3201/eid1607.090536](https://doi.org/10.3201/eid1607.090536)
- Legrand J et al. (2007) *Epidemiol Infect* 135:610. [10.1017/S0950268806007217](https://doi.org/10.1017/S0950268806007217)

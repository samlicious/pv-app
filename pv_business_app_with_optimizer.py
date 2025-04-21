"""
PV Business Case Web Application – with Credit Financing, Multi‑Currency & Optimizer
-----------------------------------------------------------------------------------
Streamlit app that replicates the Excel business‑case model and adds:
1. Debt financing with grace period finishing at fixed maturity.
2. Multi‑currency (EUR, USD, TND).
3. Optimizer that can sweep one or several parameters (debt share, interest rate, capacity) to maximise a KPI (NPV, IRR or ROI).

Install requirements:
```bash
pip install streamlit pandas numpy numpy_financial matplotlib openpyxl fpdf altair
```
Run with:
```bash
streamlit run pv_business_app.py
```
"""

import io
import itertools
import numpy as np
import pandas as pd
import streamlit as st
import numpy_financial as npf
import matplotlib.pyplot as plt
from fpdf import FPDF
import altair as alt

# ---------------------- Helper --------------------------------------------

def annuity_payment(rate: float, nper: int, pv: float) -> float:
    """Return annual annuity (positive value)"""
    return -npf.pmt(rate, nper, pv) if nper > 0 else 0.0

# ------------------- Sidebar – Base Inputs ---------------------------------

st.set_page_config(page_title="PV Business Case", layout="wide")
st.title("☀️ PV Business Case Calculator – Advanced")

sb = st.sidebar
sb.header("Project parameters")
capacity_kwp = sb.number_input("Capacity (kWp)", 1.0, 1e6, 1000.0)
project_years = sb.slider("Lifetime (yrs)", 1, 40, 20)
degradation = sb.number_input("Degradation (%/yr)", 0.0, 5.0, 0.5)/100

gti = sb.number_input("Global tilted irradiation (kWh/m²)", 500.0, 3000.0, 1800.0)
pr = sb.number_input("Performance ratio", 0.5, 0.9, 0.78)
if sb.checkbox("Override specific yield"):
    spec_yield = sb.number_input("Specific yield (kWh/kWp)", 0.0, 2500.0, 1400.0)
else:
    spec_yield = gti * pr
    sb.info(f"Specific yield = {spec_yield:.0f} kWh/kWp")

sb.header("Economics & Currency")
currency = sb.selectbox("Currency", ["EUR", "USD", "TND"], index=0)
capex = sb.number_input(f"CAPEX ({currency})", 0.0, 1e9, 800_000.0)
opex0 = sb.number_input(f"OPEX Yr1 ({currency})", 0.0, 1e8, 15_000.0)
opex_esc = sb.number_input("OPEX escalation (%/yr)", 0.0, 20.0, 2.0)/100
# Price input with 3 decimals for TND
if currency == "TND":
    price0 = sb.number_input(
        f"Price ({currency}/kWh)", 0.0, 10.0, 0.070,
        step=0.001, format="%.3f"
    )
else:
    price0 = sb.number_input(
        f"Price ({currency}/kWh)", 0.0, 10.0, 0.07,
        step=0.001
    )
price_esc = sb.number_input("Price escalation (%/yr)", 0.0, 20.0, 2.0)/100
discount = sb.number_input("Discount rate (%/yr)", 0.0, 20.0, 6.0)/100
sym = {"EUR":"€","USD":"$","TND":"TND"}[currency]

sb.header("Debt financing")
debt_share = sb.slider("Debt share (%)", 0, 100, 0)/100
debt_amt = capex * debt_share
equity_amt = capex - debt_amt
sb.write(f"Debt: {debt_amt:,.0f} {sym}, Equity: {equity_amt:,.0f} {sym}")
grace = sb.number_input("Grace period (yrs)", 0, project_years, 0)
if debt_share > 0:
    int_rate = sb.number_input("Interest rate (%/yr)", 0.0, 20.0, 4.0)/100
    maturity = sb.slider("Loan maturity (yrs)", 1, project_years, project_years)
else:
    int_rate = 0.0
    maturity = project_years

# ----------------------- Core Calculation ----------------------------------

years = np.arange(project_years)
energy = spec_yield * capacity_kwp * (1 - degradation) ** years
revenue = energy * price0 * (1 + price_esc) ** years
opex = opex0 * (1 + opex_esc) ** years

debt_svc = np.zeros(project_years)
if debt_share > 0 and maturity > grace:
    debt_svc[:grace] = debt_amt * int_rate
    ann = annuity_payment(int_rate, maturity - grace, debt_amt)
    debt_svc[grace:maturity] = ann

net_cash = revenue - opex - debt_svc
cf = np.concatenate(([-equity_amt], net_cash))
cum_cf = np.cumsum(cf)
payback = next((i for i, v in enumerate(cum_cf) if v >= 0), None)
npv_val = npf.npv(discount, cf)
irr_val = npf.irr(cf)
roi_val = cum_cf[-1] / equity_amt if equity_amt else np.nan

# -------------------------- Outputs ----------------------------------------
st.header("KPIs (Equity view)")
cols = st.columns(6)
cols[0].metric("Yield Y1 (MWh)", f"{energy[0]/1000:.1f}")
cols[1].metric("Revenue Y1", f"{revenue[0]:,.0f} {sym}")
cols[2].metric("Payback (yrs)", payback if payback is not None else "N/A")
cols[3].metric("ROI (%)", f"{roi_val*100:.1f}%")
cols[4].metric("NPV", f"{npv_val:,.0f} {sym}")
cols[5].metric("IRR (%)", "N/A" if np.isnan(irr_val) else f"{irr_val*100:.1f}%")

# Cumulative cash flow chart
st.bar_chart(pd.DataFrame({'Cumulative': cum_cf}, index=np.arange(project_years+1)))

# Cash flow summary table
cf_df = pd.DataFrame({
    'Year': np.arange(len(cf)),
    'Cash flow': cf,
    'Cumulative CF': cum_cf
})
st.write("### Cash Flow Details")
st.dataframe(cf_df.style.format({
    'Cash flow': '{:,.0f}',
    'Cumulative CF': '{:,.0f}'
}))

# -------------------------- Downloads --------------------------------------

def to_xl(sheets: dict) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for name, df in sheets.items():
            df.to_excel(w, sheet_name=name, index=False)
    return buf.getvalue()

def to_pdf(df: pd.DataFrame) -> bytes:
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font('Arial', size=12)
    pdf.cell(0, 10, 'PV Business Case - Summary', ln=True, align='C')
    pdf.ln(5)
    pdf.cell(0, 8, f"NPV: {npv_val:,.0f} {currency}", ln=True)
    pdf.cell(0, 8, f"IRR: {'N/A' if np.isnan(irr_val) else irr_val*100:.1f}%", ln=True)
    pdf.cell(0, 8, f"ROI: {roi_val*100:.1f}%", ln=True)
    pdf.cell(0, 8, f"Payback: {payback}", ln=True)
    return pdf.output(dest='S').encode('latin-1', 'replace')

st.download_button("Download Excel", data=to_xl({'CashFlow': pd.DataFrame(cf), 'Summary': pd.DataFrame(cum_cf)}), file_name="pv.xlsx")
st.download_button("Download PDF", data=to_pdf(pd.DataFrame(cf)), file_name="pv.pdf")

# -------------------------- Optimizer --------------------------------------

sb.subheader("Optimizer")
opt_kpi = sb.selectbox("KPI to maximise", ["NPV", "IRR", "ROI"])
params = sb.multiselect("Parameters to sweep", ["Debt share", "Interest rate", "Capacity"])

ranges = {}
for p in params:
    if p == "Debt share":
        low, high = sb.slider("Debt share range (%)", 0, 100, (int(debt_share*100), max(1, int(debt_share*100))))
        step = sb.number_input("Debt share step (%)", 1, 100, 5)
        ranges[p] = np.arange(low, high+step, step)/100
    elif p == "Interest rate":
        low, high = sb.slider("Interest rate range (%)", 0.0, 20.0, (int_rate*100, max(int_rate*100, 0.1)), step=0.1)
        step = sb.number_input("Interest rate step (pp)", 0.1, 5.0, 1.0)
        ranges[p] = np.arange(low, high+step, step)  # percent units
    elif p == "Capacity":
        low, high = sb.slider("Capacity range (kWp)", 1000.0, 10000.0, (capacity_kwp, capacity_kwp))
        step = sb.number_input("Capacity step (kWp)", 50.0, 5000.0, 100.0)
        ranges[p] = np.arange(low, high+step, step)


def optimise():
    if not ranges:
        st.info("Select parameters and ranges first."); return
    keys, lists = zip(*ranges.items())
    grid = [dict(zip(keys, vals)) for vals in itertools.product(*lists)]
    records = []
    for combo in grid:
        ds = combo.get('Debt share', debt_share)
        ir_pct = combo.get('Interest rate', int_rate*100)
        cap = combo.get('Capacity', capacity_kwp)
        ir = ir_pct/100
        # skip invalid maturity
        if maturity <= grace: continue
        # recalc
        ene = spec_yield * cap * (1-degradation)**years
        rev = ene * price0 * (1+price_esc)**years
        opx = opex0*(1+opex_esc)**years
        da  = capex*ds
        ds_arr = np.zeros(project_years)
        if ds>0:
            ds_arr[:grace] = da*ir
            ds_arr[grace:maturity] = annuity_payment(ir, maturity-grace, da)
        cf_loc = np.concatenate(([-(capex-da)], rev-opx-ds_arr))
        cum = np.cumsum(cf_loc)
        npv = npf.npv(discount, cf_loc)
        irr = npf.irr(cf_loc)
        roi = cum[-1] / (capex-da) if capex!=da else np.nan
        kpi = npv if opt_kpi=='NPV' else (irr if opt_kpi=='IRR' else roi)
        records.append({**combo, opt_kpi: kpi})

    df_opt = pd.DataFrame(records)
    st.subheader("Optimization results")
    if len(keys)==1:
        st.line_chart(df_opt.set_index(keys[0])[opt_kpi])
    elif len(keys)==2:
        k0,k1 = keys
        chart = alt.Chart(df_opt).mark_rect().encode(
            x=alt.X(f'{k1}:O', title=k1),
            y=alt.Y(f'{k0}:O', title=k0),
            color=alt.Color(f'{opt_kpi}:Q', title=opt_kpi)
        )
        st.altair_chart(chart, use_container_width=True)
    st.dataframe(df_opt.sort_values(opt_kpi, ascending=False).head(10))

if sb.button("Run optimization"):
    optimise()

# ------------------------- End ------------------------------------------------

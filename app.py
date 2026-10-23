from pathlib import Path
import math
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from model_bridge import recalculate_workbook, load_workbook_values

st.set_page_config(page_title="Arclight | Investor Portal", page_icon="☀️", layout="wide")
BASE_DIR = Path(__file__).resolve().parent
SOURCE_FILE = BASE_DIR / "Vyomaa_P3_P4_Capital_Raise_Model_v8(3).xlsx"

st.markdown("""
<style>
.block-container{max-width:1450px;padding-top:1rem;padding-bottom:2rem}
#MainMenu,footer{visibility:hidden}
.hero{background:linear-gradient(120deg,#10243a,#1e5270);padding:26px 30px;border-radius:18px;color:#fff;margin-bottom:20px}
.hero .eyebrow{font-size:11px;letter-spacing:2px;text-transform:uppercase;color:#b8d8e9;font-weight:700}
.hero h1{font-size:32px;color:#fff;margin:8px 0 5px 0;font-weight:750}
.hero p{color:#dbe9f1;font-size:14px;margin:0}
.section-title{font-size:21px;font-weight:750;color:#172638;margin:16px 0 3px}
.section-sub{font-size:13px;color:#728197;margin:0 0 14px}
.kpi{background:white;border:1px solid #e4eaf1;border-radius:15px;padding:17px;min-height:108px;box-shadow:0 3px 10px rgba(20,40,65,.035)}
.kpi-label{font-size:12px;color:#65758a;font-weight:650}
.kpi-value{font-size:26px;color:#15273b;font-weight:750;margin-top:8px}
.kpi-note{font-size:11px;color:#8492a6;margin-top:5px}
.panel{background:#fff;border:1px solid #e4eaf1;border-radius:16px;padding:18px 20px;margin-bottom:14px}
.pill{display:inline-block;background:#e9f7f4;color:#087e70;border-radius:20px;padding:5px 10px;font-size:11px;font-weight:700}
div[data-testid="stMetric"]{background:white;border:1px solid #e4eaf1;padding:13px 15px;border-radius:14px}
[data-testid="stSidebar"]{background:#f6f8fb;border-right:1px solid #e4eaf1}
div[data-testid="stDataFrame"]{border:1px solid #e4eaf1;border-radius:12px}
</style>
""", unsafe_allow_html=True)

def n(v):
    try:
        x=float(v)
        return x if math.isfinite(x) else None
    except (TypeError,ValueError): return None

def cr(v):
    x=n(v); return "—" if x is None else f"₹{x:,.1f} Cr"

def pct(v):
    x=n(v)
    if x is None:return "—"
    return f"{x*100:.1f}%" if abs(x)<=1.5 else f"{x:.1f}%"

def mult(v):
    x=n(v); return "—" if x is None else f"{x:.2f}x"

def v(wb,s,c):
    try:return wb[s][c].value
    except Exception:return None

def sheet_frame(wb,name,max_rows=400,max_cols=35):
    ws=wb[name]
    rows=list(ws.iter_rows(max_row=min(ws.max_row,max_rows),max_col=min(ws.max_column,max_cols),values_only=True))
    return pd.DataFrame(rows).dropna(how="all").dropna(axis=1,how="all")

def get_overrides(scenario,custom):
    base=load_workbook_values(SOURCE_FILE)
    ws=base["Inputs"]
    if scenario=="Upside case": t,y,o,c,r,g=.03,.02,-.03,-.03,10,.70
    elif scenario=="Downside case": t,y,o,c,r,g=-.05,-.064,.10,.10,11.5,.60
    elif scenario=="Custom case": t,y,o,c,r,g=custom["tariff"]/100,custom["yield"]/100,custom["opex"]/100,custom["capex"]/100,custom["rate"],custom["gearing"]/100
    else: t,y,o,c,r,g=0,0,0,0,10.5,.70
    overrides={}
    for col in ["C","D"]:
        for row,chg in [(16,t),(14,y),(18,o),(10,c)]:
            baseval=n(ws[f"{col}{row}"].value)
            if baseval is None:
                base.close(); raise ValueError(f"Inputs!{col}{row} is blank/non-numeric.")
            overrides[f"{col}{row}"]=baseval*(1+chg)
        overrides[f"{col}30"]=r/100
        overrides[f"{col}28"]=g
    overrides["C103"]=0
    base.close()
    return overrides

with st.sidebar:
    st.markdown("## ☀️ Arclight")
    st.caption("INVESTOR PORTAL")
    page=st.radio("NAVIGATION",[
        "Overview","Portfolio performance","Capital structure","Scenario lab","Model details","Integrity checks"
    ],label_visibility="collapsed")
    st.divider()
    scenario=st.selectbox("SCENARIO",["Base case","Upside case","Downside case","Custom case"],key="scenario")
    custom={}
    if scenario=="Custom case":
        custom["tariff"]=st.slider("PPA tariff change (%)",-15.0,15.0,0.0,.5)
        custom["yield"]=st.slider("Generation / yield change (%)",-15.0,10.0,0.0,.5)
        custom["opex"]=st.slider("Operating cost change (%)",-10.0,30.0,0.0,.5)
        custom["capex"]=st.slider("EPC CapEx change (%)",-10.0,30.0,0.0,.5)
        custom["rate"]=st.slider("Senior interest rate (%)",7.0,16.0,10.5,.25)
        custom["gearing"]=st.slider("Maximum gearing (%)",40,75,70,1)
    apply=st.button("Apply scenario",type="primary",use_container_width=True)
    st.divider()
    st.caption("Confidential • Development build")

if "path" not in st.session_state: st.session_state.path=None
if "done_scenario" not in st.session_state: st.session_state.done_scenario=None
if apply or not st.session_state.path or not Path(str(st.session_state.path)).exists():
    try:
        overrides=get_overrides(scenario,custom)
        with st.spinner(f"Recalculating {scenario}..."):
            path=recalculate_workbook(scenario=scenario,overrides=overrides)
        st.session_state.path=str(path); st.session_state.done_scenario=scenario
    except Exception as e:
        st.error(f"Could not recalculate workbook: {e}"); st.stop()

if st.session_state.done_scenario!=scenario:
    st.warning(f"Displayed values are for {st.session_state.done_scenario}. Click Apply scenario to load {scenario}.")

wb=load_workbook_values(st.session_state.path)
capacity=n(v(wb,"Dashboard","E5")); uses=n(v(wb,"Dashboard","E7")); debt=n(v(wb,"Dashboard","E8"))
raise_amt=n(v(wb,"Dashboard","E29")); irr=n(v(wb,"Dashboard","E35")); moic=n(v(wb,"Dashboard","E36"))
p3p50=n(v(wb,"Dashboard","C17"));p4p50=n(v(wb,"Dashboard","D17"));p3p90=n(v(wb,"Dashboard","C18"));p4p90=n(v(wb,"Dashboard","D18"))

st.markdown(f"""<div class="hero"><div class="eyebrow">Investment workspace &nbsp; • &nbsp; {st.session_state.done_scenario}</div><h1>Arclight Renewables</h1><p>Commercial & institutional solar | P3 / P4 portfolio financing and growth</p></div>""",unsafe_allow_html=True)

def kpi(label,value,note):
    st.markdown(f'<div class="kpi"><div class="kpi-label">{label}</div><div class="kpi-value">{value}</div><div class="kpi-note">{note}</div></div>',unsafe_allow_html=True)

def section(title,subtitle=""):
    st.markdown(f'<div class="section-title">{title}</div><div class="section-sub">{subtitle}</div>',unsafe_allow_html=True)

# Portfolio performance is now a visual page; no raw Excel table is shown by default.
if page=="Overview":
    section("Investment snapshot","Key model outputs at a glance")
    cols=st.columns(4)
    with cols[0]: kpi("Portfolio capacity","—" if capacity is None else f"{capacity:.1f} MWp","P3 + P4")
    with cols[1]: kpi("Total project uses",cr(uses),"Capital deployment")
    with cols[2]: kpi("Senior debt",cr(debt),"Project-level financing")
    with cols[3]: kpi("Institutional equity",cr(raise_amt),"HoldCo raise")
    left,right=st.columns(2)
    with left:
        section("Capital mix","Senior debt versus HoldCo funding")
        debt_src=n(v(wb,"Sources & Uses","E29")) or 0
        equity_src=n(v(wb,"Sources & Uses","E30")) or 0
        if debt_src+equity_src:
            fig=px.pie(pd.DataFrame({"Source":["Senior debt","HoldCo injection"],"₹ Cr":[debt_src,equity_src]}),names="Source",values="₹ Cr",hole=.68)
            fig.update_traces(textposition="inside",textinfo="percent+label")
            fig.update_layout(height=330,margin=dict(t=10,b=10,l=10,r=10),showlegend=False)
            st.plotly_chart(fig,use_container_width=True)
        else: st.info("No values in mapped funding cells.")
    with right:
        section("Debt coverage","Minimum P50 and P90 DSCR")
        df=pd.DataFrame({"Portfolio":["P3","P4","P3","P4"],"Case":["P50","P50","P90","P90"],"DSCR":[p3p50,p4p50,p3p90,p4p90]}).dropna()
        if not df.empty:
            fig=px.bar(df,x="Portfolio",y="DSCR",color="Case",barmode="group",text_auto=".2f")
            fig.update_layout(height=330,margin=dict(t=10,b=10,l=10,r=10),yaxis_title="DSCR (x)",legend_title="")
            st.plotly_chart(fig,use_container_width=True)
        else: st.info("DSCR cells are blank.")
    section("Investor returns","Indicative outputs from HoldCo model")
    cols=st.columns(3)
    with cols[0]: kpi("Partner IRR",pct(irr),"Model output • not guaranteed")
    with cols[1]: kpi("Partner MOIC",mult(moic),"Model output")
    with cols[2]: kpi("Active scenario",st.session_state.done_scenario,"Recalculated case")

elif page=="Portfolio performance":
    section("Portfolio performance","Visual forecast analysis for SPV3 (P3) and SPV4 (P4)")
    chosen=st.radio("Portfolio view",["Combined","P3 only","P4 only"],horizontal=True)
    sheets=["SPV3","SPV4"] if chosen=="Combined" else (["SPV3"] if chosen=="P3 only" else ["SPV4"])
    rows=[]
    for sh in sheets:
        ws=wb[sh]
        for col in range(5,min(ws.max_column,35)+1):
            year=n(ws.cell(6,col).value)
            revenue=n(ws.cell(13,col).value); ebitda=n(ws.cell(16,col).value)
            debt_service=n(ws.cell(30,col).value); dscr=n(ws.cell(31,col).value)
            if year is not None and any(x is not None for x in [revenue,ebitda]):
                rows.append({"Year":int(year),"Portfolio":sh,"Revenue":revenue,"EBITDA":ebitda,"Debt service":debt_service,"DSCR":dscr})
    df=pd.DataFrame(rows)
    if df.empty:
        st.warning("No forecast data found using the current SPV3/SPV4 cell mapping.")
    else:
        # Summary cards derived from available series
        last=df.sort_values("Year").groupby("Portfolio",as_index=False).tail(1)
        revenue_latest=last["Revenue"].sum(min_count=1)
        ebitda_latest=last["EBITDA"].sum(min_count=1)
        c1,c2,c3=st.columns(3)
        with c1:kpi("Latest forecast revenue",cr(revenue_latest),"Last forecast year shown")
        with c2:kpi("Latest forecast EBITDA",cr(ebitda_latest),"Last forecast year shown")
        with c3:kpi("Forecast period",f"{df['Year'].min():.0f}–{df['Year'].max():.0f}","Financial years")
        section("Revenue and EBITDA trend","Compare portfolio operating performance over time")
        fig=go.Figure()
        for sh in df["Portfolio"].unique():
            sub=df[df["Portfolio"]==sh]
            fig.add_trace(go.Scatter(x=sub["Year"],y=sub["Revenue"],mode="lines+markers",name=f"{sh} revenue",line=dict(width=3)))
            fig.add_trace(go.Scatter(x=sub["Year"],y=sub["EBITDA"],mode="lines+markers",name=f"{sh} EBITDA",line=dict(width=2,dash="dot")))
        fig.update_layout(height=420,hovermode="x unified",margin=dict(t=20,b=20,l=10,r=10),legend=dict(orientation="h",y=-.22),yaxis_title="₹ Cr",xaxis_title="")
        st.plotly_chart(fig,use_container_width=True)
        left,right=st.columns(2)
        with left:
            section("Debt service coverage","Annual DSCR trajectory")
            dscr_df=df.dropna(subset=["DSCR"])
            if not dscr_df.empty:
                fig=px.line(dscr_df,x="Year",y="DSCR",color="Portfolio",markers=True)
                fig.add_hline(y=1.0,line_dash="dash",annotation_text="1.00x threshold")
                fig.update_layout(height=340,margin=dict(t=15,b=15,l=10,r=10),yaxis_title="DSCR (x)")
                st.plotly_chart(fig,use_container_width=True)
            else: st.info("No DSCR values found in the mapped row.")
        with right:
            section("Debt service profile","Annual scheduled debt service")
            ds_df=df.dropna(subset=["Debt service"])
            if not ds_df.empty:
                fig=px.area(ds_df,x="Year",y="Debt service",color="Portfolio")
                fig.update_layout(height=340,margin=dict(t=15,b=15,l=10,r=10),yaxis_title="₹ Cr")
                st.plotly_chart(fig,use_container_width=True)
            else: st.info("No debt service values found in the mapped row.")
        with st.expander("View underlying forecast data"):
            st.dataframe(df,width="stretch",hide_index=True)

elif page=="Capital structure":
    section("Capital structure","Funding stack and investor economics")
    cols=st.columns(4)
    with cols[0]:kpi("Total uses",cr(uses),"Capital requirement")
    with cols[1]:kpi("Senior debt",cr(debt),"Project finance")
    with cols[2]:kpi("Institutional equity",cr(raise_amt),"HoldCo capital")
    with cols[3]:kpi("Partner IRR",pct(irr),"Indicative model output")
    section("Sources & uses","Expand only when you need detailed schedules")
    with st.expander("View Sources & Uses schedule"):
        st.dataframe(sheet_frame(wb,"Sources & Uses",180,20),width="stretch",hide_index=True)
    with st.expander("View Funding Plan"):
        st.dataframe(sheet_frame(wb,"Funding Plan",180,20),width="stretch",hide_index=True)
    with st.expander("View HoldCo return waterfall"):
        st.dataframe(sheet_frame(wb,"HoldCo",180,25),width="stretch",hide_index=True)

elif page=="Scenario lab":
    section("Scenario lab","Change assumptions in the sidebar and compare the recalculated outputs")
    c1,c2,c3,c4=st.columns(4)
    with c1:kpi("Current case",st.session_state.done_scenario,"Applied workbook")
    with c2:kpi("Total uses",cr(uses),"Current scenario")
    with c3:kpi("Senior debt",cr(debt),"Current scenario")
    with c4:kpi("Partner IRR",pct(irr),"Current scenario")
    st.info("Choose Base, Upside, Downside or Custom in the sidebar, then click Apply scenario. The source workbook is not modified.")
    with st.expander("View assumptions worksheet"):
        st.dataframe(sheet_frame(wb,"Inputs",180,12),width="stretch",hide_index=True)

elif page=="Model details":
    section("Model details","Detailed schedules are available on demand, not as the default view")
    sh=st.selectbox("Choose a worksheet",wb.sheetnames,index=wb.sheetnames.index("Dashboard") if "Dashboard" in wb.sheetnames else 0)
    with st.expander("Open worksheet data",expanded=False):
        st.dataframe(sheet_frame(wb,sh,500,40),width="stretch",height=600,hide_index=True)
    st.download_button("Download calculated workbook",Path(st.session_state.path).read_bytes(),file_name=f"Arclight_{(st.session_state.done_scenario or 'scenario').replace(' ','_')}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif page=="Integrity checks":
    section("Model integrity","Review check outcomes in a readable view")
    checkws=wb["Checks"]
    records=[]
    for r in range(1,checkws.max_row+1):
        label=checkws.cell(r,2).value
        result=checkws.cell(r,3).value
        if label is not None:
            status="PASS" if result is True else ("FAIL" if result is False else ("REVIEW" if result is None else str(result)))
            records.append({"Check":str(label),"Status":status})
    checkdf=pd.DataFrame(records)
    if not checkdf.empty:
        total=len(checkdf); failed=(checkdf["Status"]=="FAIL").sum(); passed=(checkdf["Status"]=="PASS").sum()
        a,b,c=st.columns(3)
        with a:kpi("Checks reviewed",str(total),"Workbook check items")
        with b:kpi("Passed",str(passed),"Calculated TRUE")
        with c:kpi("Failed",str(failed),"Calculated FALSE")
        for _,row in checkdf.iterrows():
            icon="✅" if row["Status"]=="PASS" else ("❌" if row["Status"]=="FAIL" else "⚪")
            st.markdown(f'<div class="panel" style="padding:12px 16px;margin-bottom:8px"><b>{icon} &nbsp; {row["Check"]}</b><span style="float:right;color:#64748b">{row["Status"]}</span></div>',unsafe_allow_html=True)
    else:
        st.info("No check labels found in column B of the Checks worksheet.")

wb.close()
st.markdown('<div style="color:#8794a6;font-size:11px;margin-top:26px">CONFIDENTIAL • DEVELOPMENT BUILD • Validate all calculations against Excel before investor circulation.</div>',unsafe_allow_html=True)

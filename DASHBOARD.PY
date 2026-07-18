import pandas as pd
import streamlit as st
import numpy as np
import plotly.express as px
import io
from datetime import datetime, timedelta

# Page Configuration
st.set_page_config(page_title="Tertiary Sale Dashboard", layout="wide")
st.title("Tertiary Sale Dashboard 📊")

# ==========================================
# 📆 GLOBAL SCOPE: DYNAMIC TIME & SEQUENCE LOGIC
# ==========================================
today = datetime.now()
target_date = today - timedelta(days=1)
current_month_num = target_date.month
current_day_limit = target_date.day
current_date_str = target_date.strftime("%d-%b")

passed_months = [m for m in range(4, 13) if m < current_month_num] if current_month_num >= 4 else [m for m in range(4, 13)] + [m for m in range(1, current_month_num)]

m_sequence = ['April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December', 'January', 'February', 'March']

months_to_show = []
temp_date = datetime(2025, 4, 1)
while len(months_to_show) < 4:
    if temp_date.month == current_month_num:
         months_to_show.append(temp_date.strftime('%B'))
         break
    months_to_show.append(temp_date.strftime('%B'))
    temp_date = (temp_date.replace(day=1) + timedelta(days=32)).replace(day=1)
    if len(months_to_show) == 4: break

# Custom CSS for Tabs 
st.markdown("""
<style>
    button[data-baseweb="tab"] {
        border: 2px solid #e0e0e0 !important;
        border-radius: 8px !important;
        padding: 10px 16px !important;
        margin: 5px !important;
        background-color: #f8f9fa !important;
        min-width: 140px !important;
        font-weight: bold !important;
        font-size: 15px !important;
    }
    button[data-baseweb="tab"][aria-selected="true"] {
        background-color: #1F4E79 !important;
        color: white !important;
        border-color: #1F4E79 !important;
    }
    div[data-testid="stTabs"] div[role="tablist"] {
        display: flex !important;
        flex-wrap: wrap !important;
        overflow-x: hidden !important;
    }
</style>
""", unsafe_allow_html=True)

# Download Function
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')
    return output.getvalue()

# Premium Indian Lakhs/Crores Comma Formatting Function
def format_indian_num(val, is_currency=False, prefix="₹"):
    if not isinstance(val, (int, float, np.number)) or pd.isna(val):
        return str(val)
    is_negative = val < 0
    val = abs(val)
    s = f"{val:.0f}"
    if len(s) <= 3:
        res = s
    else:
        res = s[-3:]
        remaining = s[:-3]
        while remaining:
            res = remaining[-2:] + "," + res
            remaining = remaining[:-2]
    if is_negative:
        res = "-" + res
    return f"{prefix}{res}" if is_currency else res

@st.cache_data(ttl=3600)
def load_data():
    try:
        df = pd.read_excel('TERTIARY_SALE.xlsx', engine='openpyxl')
    except FileNotFoundError:
        st.error("⚠️ TERTIARY_SALE.xlsx nahi mili.")
        st.stop()
    df.columns = df.columns.str.strip().str.upper()
    df['SALEVAL'] = pd.to_numeric(df['SALEVAL'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
    df['QTY'] = pd.to_numeric(df['QTY'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
    df['BILLDATE'] = pd.to_datetime(df['BILLDATE'], errors='coerce').fillna(pd.to_datetime('2000-01-01'))
    df['MONTH_NAME'] = df['BILLDATE'].dt.strftime('%B')
    df['YEAR2'] = df['YEAR2'].astype(str)
    
    text_cols = ['FORMAL/CASUAL', 'EOSS/FRESH', 'SEASON', 'FIT', 'CATEGORY', 'MANAGER', 'STATE', 'STATUS', 'CUSTOMERNAMEACTIVE']
    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).replace(r'[\xa0\t\n\r]', ' ', regex=True)
            df[col] = df[col].str.replace(r'\s+', ' ', regex=True)
            df[col] = df[col].str.strip().str.upper()
            
    if 'CUSTOMERNAMEACTIVE' in df.columns:
        mask_anomaly = (df['CUSTOMERNAMEACTIVE'].str.contains('MOM', na=False)) & \
                       (df['BILLDATE'].dt.year == 2026) & \
                       (df['BILLDATE'].dt.month.isin([8, 9, 10, 11, 12]))
        df.loc[mask_anomaly, 'SALEVAL'] = 0
        df.loc[mask_anomaly, 'QTY'] = 0
            
    return df

def style_dataframe(df):
    sale_props = {'background-color': '#d4edda', 'border': '1px solid #28a745'}
    growth_props = {'background-color': '#ebf7ed', 'border': '1px solid #a9d08e'}
    styler = df.style.hide(axis="index")
    for col in df.columns:
        if 'GR(%)' in col.upper():
            styler = styler.set_properties(**growth_props, subset=[col])
        elif 'SALE' in col.upper() or ('LEDGER' in col.upper() and 'QTY' not in col.upper()):
            styler = styler.set_properties(**sale_props, subset=[col])
    return styler

def calc_growth(base, comp):
    if base == 0:
        if comp == 0: return 0.0
        return np.inf if comp > 0 else np.nan
    return (comp - base) / base * 100

def calc_growth_series(base_s, comp_s):
    base_s = base_s.astype(float)
    comp_s = comp_s.astype(float)
    with np.errstate(divide='ignore', invalid='ignore'):
        result = (comp_s - base_s) / base_s.replace(0, np.nan) * 100
    zero_base = base_s == 0
    result = result.mask(zero_base & (comp_s == 0), 0.0)
    result = result.mask(zero_base & (comp_s > 0), np.inf)
    return result

def format_cell(x, col_name):
    if not isinstance(x, (int, float, np.number)): return x
    if "Gr(%)" in col_name:
        if x == np.inf: return "🆕 New"
        if pd.isna(x): return "N/A"
        return f"📈 🟢 {x:.0f}%" if x > 0 else f"📉 🔴 {abs(x):.0f}%" if x < 0 else "➖ 0%"
    if pd.isna(x): return "N/A"
    if "SALE" in col_name.upper() or ("LEDGER" in col_name.upper() and "QTY" not in col_name.upper()):
        return format_indian_num(x, is_currency=True)
    return format_indian_num(x, is_currency=False)

sheet_data = load_data()

# ==========================================
# 🔍 SIDEBAR FILTERS FORM
# ==========================================
st.sidebar.markdown("### 🎯 Global Filters")

with st.sidebar.form(key='global_filter_form'):
    fy_list = sorted(sheet_data['YEAR2'].unique())
    col_y1, col_y2 = st.columns(2)
    base_y = col_y1.selectbox("Base Year (LY):", fy_list, index=max(0, len(fy_list)-2))
    comp_y = col_y2.selectbox("Compare Year (TY):", fy_list, index=max(0, len(fy_list)-1))
    
    st.markdown("---")
    
    all_columns = list(sheet_data.columns)
    guess_idx = 0
    for i, c in enumerate(all_columns):
        if c in ['BILLDATE', 'MONTH_NAME', 'YEAR2']:
            continue
        if any(k in c for k in ['BILLNO', 'INVOICENO', 'BILL_NO', 'INVOICE_NO', 'BILL NO', 'INVOICE NO', 'BILL', 'INVOICE', 'NUMBER', 'INV']):
            guess_idx = i
            break
    selected_bill_col = st.selectbox("🧾 Select Bill/Invoice Column:", all_columns, index=guess_idx)
    
    st.markdown("---")
    search_cust = st.text_input("🔍 Search Customer Name:").strip().upper()
    
    filter_cols = ['MONTH_NAME', 'CUSTOMERNAMEACTIVE', 'MANAGER', 'CATEGORY', 'FIT', 'SEASON', 'STATE', 'STATUS', 'EOSS/FRESH', 'FORMAL/CASUAL']
    selected_filters = {}
    
    for col in filter_cols:
        if col in sheet_data.columns:
            unique_vals = sorted([str(x) for x in sheet_data[col].dropna().unique()])
            default_vals = None
            if col == 'CATEGORY':
                exclude_cats = ['BAGS DUFFLE BAG', 'BAGS', 'SUITCOVERS', 'TROLLEY BAG', 'DUFFLE BAG', 'BAG-LEATHERS']
                default_vals = [c for c in unique_vals if c not in exclude_cats]
                
            selected_filters[col] = st.multiselect(f"Filter by {col.replace('_', ' ').title()}:", unique_vals, default=default_vals)
            
    submit_button = st.form_submit_button(label='Apply Filters 🚀', use_container_width=True)

filtered_df = sheet_data.copy()
if search_cust:
    filtered_df = filtered_df[filtered_df['CUSTOMERNAMEACTIVE'].astype(str).str.contains(search_cust, case=False, na=False)]

for col, selected_vals in selected_filters.items():
    if len(selected_vals) > 0:
        filtered_df = filtered_df[filtered_df[col].astype(str).isin(selected_vals)]

if selected_bill_col in filtered_df.columns and 'CUSTOMERNAMEACTIVE' in filtered_df.columns:
    filtered_df['UNIQUE_BILL_ID'] = filtered_df['YEAR2'].astype(str) + "_" + filtered_df['MONTH_NAME'].astype(str) + "_" + filtered_df['CUSTOMERNAMEACTIVE'].astype(str) + "_" + filtered_df[selected_bill_col].astype(str)
else:
    filtered_df['UNIQUE_BILL_ID'] = filtered_df['YEAR2'].astype(str) + "_" + filtered_df['MONTH_NAME'].astype(str) + "_" + filtered_df['CUSTOMERNAMEACTIVE'].astype(str) + "_" + filtered_df['BILLDATE'].astype(str)

df_comp_kpi = filtered_df[filtered_df['YEAR2'] == comp_y]
df_base_kpi = filtered_df[filtered_df['YEAR2'] == base_y]

# Global KPI aggregates
s_comp = df_comp_kpi['SALEVAL'].sum()
s_base = df_base_kpi['SALEVAL'].sum()
q_comp = df_comp_kpi['QTY'].sum()
q_base = df_base_kpi['QTY'].sum()
d_comp = df_comp_kpi['CUSTOMERNAMEACTIVE'].nunique() if 'CUSTOMERNAMEACTIVE' in filtered_df.columns else 0
d_base = df_base_kpi['CUSTOMERNAMEACTIVE'].nunique() if 'CUSTOMERNAMEACTIVE' in filtered_df.columns else 0
b_comp = df_comp_kpi['UNIQUE_BILL_ID'].nunique() if 'UNIQUE_BILL_ID' in df_comp_kpi.columns else 0
b_base = df_base_kpi['UNIQUE_BILL_ID'].nunique() if 'UNIQUE_BILL_ID' in df_base_kpi.columns else 0

abv_comp = s_comp / b_comp if b_comp > 0 else 0
abv_base = s_base / b_base if b_base > 0 else 0
aqv_comp = s_comp / q_comp if q_comp > 0 else 0
aqv_base = s_base / q_base if q_base > 0 else 0
upt_comp = q_comp / b_comp if b_comp > 0 else 0
upt_base = q_base / b_base if b_base > 0 else 0

# Dynamic Year-to-Date isolation ranges
df_ytd_global = filtered_df[
    (filtered_df['BILLDATE'].dt.month.isin(passed_months)) | 
    ((filtered_df['BILLDATE'].dt.month == current_month_num) & (filtered_df['BILLDATE'].dt.day <= current_day_limit))
]
df_ytd_comp = df_ytd_global[df_ytd_global['YEAR2'] == comp_y]
df_ytd_base = df_ytd_global[df_ytd_global['YEAR2'] == base_y]

s_comp_ytd_g, s_base_ytd_g = df_ytd_comp['SALEVAL'].sum(), df_ytd_base['SALEVAL'].sum()
q_comp_ytd_g, q_base_ytd_g = df_ytd_comp['QTY'].sum(), df_ytd_base['QTY'].sum()
b_comp_ytd_g, b_base_ytd_g = df_ytd_comp['UNIQUE_BILL_ID'].nunique() if 'UNIQUE_BILL_ID' in df_ytd_comp.columns else 0, df_ytd_base['UNIQUE_BILL_ID'].nunique() if 'UNIQUE_BILL_ID' in df_ytd_base.columns else 0
d_comp_ytd_g, d_base_ytd_g = df_ytd_comp['CUSTOMERNAMEACTIVE'].nunique() if 'CUSTOMERNAMEACTIVE' in df_ytd_comp.columns else 0, df_ytd_base['CUSTOMERNAMEACTIVE'].nunique() if 'CUSTOMERNAMEACTIVE' in df_ytd_base.columns else 0

# ==========================================
# 🌟 GLOBAL PERFORMANCE AREA 
# ==========================================
st.markdown(f"### 🌟 Global Performance Summary ({comp_y} vs {base_y})")
r1_c1, r1_c2, r1_c3, r1_c4 = st.columns(4)
r1_c1.metric(label="Total Sale", value=format_indian_num(s_comp, True), delta=f"{calc_growth(s_base, s_comp):.1f}%")
r1_c2.metric(label="Total Qty", value=format_indian_num(q_comp, False), delta=f"{calc_growth(q_base, q_comp):.1f}%")
r1_c3.metric(label="Total No Bills", value=format_indian_num(b_comp, False), delta=f"{calc_growth(b_base, b_comp):.1f}%")
r1_c4.metric(label="Total Active Doors", value=format_indian_num(d_comp, False), delta=f"{calc_growth(d_base, d_comp):.1f}%")

r2_c1, r2_c2, r2_c3, r2_c4 = st.columns(4)
r2_c1.metric(label="Average Bill Value", value=format_indian_num(abv_comp, True), delta=f"{calc_growth(abv_base, abv_comp):.1f}%")
r2_c2.metric(label="Average Qty Value", value=format_indian_num(aqv_comp, True), delta=f"{calc_growth(aqv_base, aqv_comp):.1f}%")
r2_c3.metric(label="Unit Per Transaction", value=f"{upt_comp:.2f}", delta=f"{calc_growth(upt_base, upt_comp):.1f}%")

with r2_c4:
    with st.container(border=True):
        v_grow = calc_growth(s_base_ytd_g, s_comp_ytd_g)
        q_grow = calc_growth(q_base_ytd_g, q_comp_ytd_g)
        v_color = "#2E7D32" if v_grow >= 0 else "#D32F2F"
        q_color = "#2E7D32" if q_grow >= 0 else "#D32F2F"
        
        st.markdown(f"""
        <div style="font-size: 15px; font-family: Arial, sans-serif; font-weight: bold; margin-bottom: 5px;">
            📈 YTD Value Growth: <span style="color: {v_color};">{v_grow:+.1f}%</span> | Qty: <span style="color: {q_color};">{q_grow:+.1f}%</span>
        </div>
        """, unsafe_allow_html=True)
        st.download_button("📥 DOWNLOAD FULL REPORT", data=to_excel(filtered_df), file_name="Report.xlsx", mime="application/vnd.ms-excel", use_container_width=True)

st.markdown(f"### 📈 Current Growth Analytics Breakdown (Up to {current_date_str})")
with st.container(border=True):
    ytd_l1, ytd_l2, ytd_l3, ytd_l4 = st.columns(4)
    ytd_l1.metric(label="YTD Sales Value", value=format_indian_num(s_comp_ytd_g, True), delta=f"{v_grow:.1f}%")
    ytd_l2.metric(label="YTD Total Quantity", value=format_indian_num(q_comp_ytd_g, False), delta=f"{q_grow:.1f}%")
    ytd_l3.metric(label="YTD Invoice Counts", value=format_indian_num(b_comp_ytd_g, False), delta=f"{calc_growth(b_base_ytd_g, b_comp_ytd_g):.1f}%")
    ytd_l4.metric(label="YTD Active Stores", value=format_indian_num(d_comp_ytd_g, False), delta=f"{calc_growth(d_base_ytd_g, d_comp_ytd_g):.1f}%")
    
    st.markdown("##### 📅 Passed & Running Month-wise Live Sales Values")
    m_cols = st.columns(4)
    for idx, m_name in enumerate(months_to_show):
        if m_name == target_date.strftime('%B'):
            df_m_c = df_ytd_comp[df_ytd_comp['MONTH_NAME'] == m_name]
            df_m_b = df_ytd_base[df_ytd_base['MONTH_NAME'] == m_name]
        else:
            df_m_c = df_comp_kpi[df_comp_kpi['MONTH_NAME'] == m_name]
            df_m_b = df_base_kpi[df_base_kpi['MONTH_NAME'] == m_name]
        s_m_c, s_m_b = df_m_c['SALEVAL'].sum(), df_m_b['SALEVAL'].sum()
        m_cols[idx].metric(label=f"📅 {m_name} Sale", value=format_indian_num(s_m_c, True), delta=f"{calc_growth(s_m_b, s_m_c):.1f}%")

st.markdown("---")

# Navigation setup
tab_current_graph, tab_graph, tab_growth, tab1, tab_week, tab2, tab_cust, tab_cust_month, tab3, tab4, tab5 = st.tabs([
    "📅 Current Date Analytics Summary", "📊 Analytics Summary", "🚀 Store Growth Deep-Dive", "📈 YoY Growth", "📅 Week-wise Analysis",
    "👕 Attributes", "👥 Customer Performance", "📅 Customer Month Analysis", 
    "📊 Category vs Time", "🏆 Top Performers", "📔 Mega Ledger"
])

# ✨ FIXED: Margin right ko kam kiya (r=220) aur x=1.02 kiya taaki pie graph ko maximize space mile aur wo screen par BADA dikhe!
def make_professional_chart(fig, title_text, is_pie=False):
    fig.update_layout(
        title={'text': title_text, 'font': {'size': 22, 'family': 'Arial', 'weight': 'bold'}, 'x': 0.02},
        font=dict(size=16, color='#333333'),
        template='plotly_white',
        margin=dict(l=10, r=220, t=85, b=45) if is_pie else dict(l=55, r=25, t=85, b=45),
        xaxis=dict(tickfont=dict(size=14), title_font=dict(size=16, weight='bold'), tickformat='.0f'),
        yaxis=dict(tickfont=dict(size=14), title_font=dict(size=16, weight='bold'), tickformat='.0f')
    )
    if is_pie:
        fig.update_layout(
            showlegend=True,
            legend=dict(
                font=dict(size=12),
                orientation="v",         
                yanchor="middle",        
                y=0.5,                   
                xanchor="left",
                x=1.02                   
            )
        )
    else:
        fig.update_layout(legend=dict(font=dict(size=13), orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
        fig.update_traces(textposition='outside', textfont=dict(size=13, weight='bold'))
    return fig

# ✨ FIXED ENGINE: Legend me <br> tag lagakar data ko customer ke naam ke strictly niche bhej diya hai!
def render_analytics(df_source, tab_key_prefix, is_ytd=False):
    plot_df = df_source[df_source['YEAR2'].isin([base_y, comp_y])].copy()
    m_seq = months_to_show if is_ytd else m_sequence

    # Graph 1: Status Header
    if 'STATUS' in plot_df.columns:
        stat_df = plot_df[plot_df['YEAR2'] == comp_y].groupby('STATUS')['CUSTOMERNAMEACTIVE'].nunique().reset_index()
        fig_stat = px.bar(stat_df, x='STATUS', y='CUSTOMERNAMEACTIVE', color='STATUS', text_auto='.0f', color_discrete_sequence=px.colors.qualitative.Dark24)
        st.plotly_chart(make_professional_chart(fig_stat, f"1. Active Stores by Status Header ({comp_y})"), use_container_width=True, key=f"{tab_key_prefix}_status_top")
    
    # 🎯 2ND SPOT PIES WITH DATA UNDER THE NAME IN LEGEND TO EXPAND PIE SIZE
    cust_kpi = plot_df[plot_df['YEAR2'] == comp_y]
    cust_sum_df = cust_kpi.groupby('CUSTOMERNAMEACTIVE').agg(VALUE=('SALEVAL', 'sum'), QTY=('QTY', 'sum')).reset_index()
    s_tot = cust_sum_df['VALUE'].sum()
    
    def build_pie_mappings(df_segment):
        inside_lbls = []
        legend_lbls = []
        for _, r in df_segment.iterrows():
            p_pct = (r['VALUE'] / s_tot * 100) if s_tot > 0 else 0
            v_str = format_indian_num(r['VALUE'], False)
            q_str = format_indian_num(r['QTY'], False)
            
            # Slices ke andar teeno main performance indicator visible rahenge
            inside_lbls.append(f"₹{v_str}<br>{q_str} Qty<br>({p_pct:.0f}%)")
            
            # ✨ Cstumer name ke sath data ko break (<br>) karke uske thik NICHE wali line me set kiya hai!
            legend_lbls.append(f"<b>{r['CUSTOMERNAMEACTIVE']}</b><br>₹{v_str} | {q_str} Qty | {p_pct:.0f}%")
        return inside_lbls, legend_lbls

    c_p1, c_p2 = st.columns(2)
    cust_top_sort = cust_sum_df.sort_values(by='VALUE', ascending=False).reset_index(drop=True)
    with c_p1:
        if not cust_top_sort.head(5).empty and s_tot > 0:
            p1_df = cust_top_sort.head(5).copy()
            in_t, leg_t = build_pie_mappings(p1_df)
            p1_df['LEGEND_TEXT'] = leg_t
            fig_p1 = px.pie(p1_df, values='VALUE', names='LEGEND_TEXT', color_discrete_sequence=px.colors.qualitative.Pastel)
            fig_p1.update_traces(text=in_t, textinfo='text', textposition='inside', textfont=dict(size=13, weight='bold'))
            st.plotly_chart(make_professional_chart(fig_p1, "2A. Top 5 High Performing Customers Share", is_pie=True), use_container_width=True, key=f"{tab_key_prefix}_p1")
    with c_p2:
        if len(cust_top_sort) > 5 and s_tot > 0:
            p2_df = cust_top_sort.iloc[5:10].copy()
            in_t, leg_t = build_pie_mappings(p2_df)
            p2_df['LEGEND_TEXT'] = leg_t
            fig_p2 = px.pie(p2_df, values='VALUE', names='LEGEND_TEXT', color_discrete_sequence=px.colors.qualitative.Set3)
            fig_p2.update_traces(text=in_t, textinfo='text', textposition='inside', textfont=dict(size=13, weight='bold'))
            st.plotly_chart(make_professional_chart(fig_p2, "2B. Rank 6-10 Performing Customers Share", is_pie=True), use_container_width=True, key=f"{tab_key_prefix}_p2")
            
    c_b1, c_b2 = st.columns(2)
    cust_bot_sort = cust_sum_df.sort_values(by='VALUE', ascending=True).reset_index(drop=True)
    
    red_palette_1 = ['#E53935', '#EF5350', '#E57373', '#EF9A9A', '#FFCDD2']
    red_palette_2 = ['#B71C1C', '#C62828', '#D32F2F', '#E53935', '#EF5350']

    with c_b1:
        if not cust_bot_sort.head(5).empty and s_tot > 0:
            b1_df = cust_bot_sort.head(5).copy()
            in_t, leg_t = build_pie_mappings(b1_df)
            b1_df['LEGEND_TEXT'] = leg_t
            fig_b1 = px.pie(b1_df, values='VALUE', names='LEGEND_TEXT', color_discrete_sequence=red_palette_1)
            fig_b1.update_traces(text=in_t, textinfo='text', textposition='inside', textfont=dict(size=13, weight='bold'))
            st.plotly_chart(make_professional_chart(fig_b1, "2C. Bottom 5 Lowest Performing Stores Share", is_pie=True), use_container_width=True, key=f"{tab_key_prefix}_b1")
    with c_b2:
        if len(cust_bot_sort) > 5 and s_tot > 0:
            b2_df = cust_bot_sort.iloc[5:10].copy()
            in_t, leg_t = build_pie_mappings(b2_df)
            b2_df['LEGEND_TEXT'] = leg_t
            fig_b2 = px.pie(b2_df, values='VALUE', names='LEGEND_TEXT', color_discrete_sequence=red_palette_2)
            fig_b2.update_traces(text=in_t, textinfo='text', textposition='inside', textfont=dict(size=13, weight='bold'))
            st.plotly_chart(make_professional_chart(fig_b2, "2D. Rank Bottom 6-10 Lowest Stores Share", is_pie=True), use_container_width=True, key=f"{tab_key_prefix}_b2")

    # Graph 3: Month wise trends
    c_m1, c_m2 = st.columns(2)
    with c_m1:
        m_v_df = plot_df.groupby(['MONTH_NAME', 'YEAR2'])['SALEVAL'].sum().reset_index()
        m_v_df['TEXT'] = m_v_df['SALEVAL'].apply(lambda x: format_indian_num(x, True))
        fig_m_v = px.bar(m_v_df, x='MONTH_NAME', y='SALEVAL', color='YEAR2', text='TEXT', barmode='group', color_discrete_sequence=['#A6C8E0', '#1F4E79'])
        fig_m_v.update_layout(xaxis={'categoryorder':'array', 'categoryarray': m_seq})
        st.plotly_chart(make_professional_chart(fig_m_v, "3A. Month-wise Revenue Value Trend (LY vs TY)"), use_container_width=True, key=f"{tab_key_prefix}_m_v")
    with c_m2:
        m_q_df = plot_df.groupby(['MONTH_NAME', 'YEAR2'])['QTY'].sum().reset_index()
        m_q_df['TEXT'] = m_q_df['QTY'].apply(lambda x: format_indian_num(x, False))
        fig_m_q = px.bar(m_q_df, x='MONTH_NAME', y='QTY', color='YEAR2', text='TEXT', barmode='group', color_discrete_sequence=['#DCA6C8', '#7030A0'])
        fig_m_q.update_layout(xaxis={'categoryorder':'array', 'categoryarray': m_seq})
        st.plotly_chart(make_professional_chart(fig_m_q, "3B. Month-wise Total Quantity Volume Trend (LY vs TY)"), use_container_width=True, key=f"{tab_key_prefix}_m_q")

    # Dynamic side by side dimensions execution loops
    def render_dimension_twins(col_name, title_num):
        if col_name in plot_df.columns:
            cx1, cx2 = st.columns(2)
            with cx1:
                if col_name == 'SEASON':
                    df_v = plot_df[plot_df['YEAR2'] == comp_y].groupby(col_name)['SALEVAL'].sum().reset_index()
                    df_v['TEXT'] = df_v['SALEVAL'].apply(lambda x: format_indian_num(x, True))
                    fig_v = px.bar(df_v, x=col_name, y='SALEVAL', text='TEXT', color_discrete_sequence=['#385723'])
                else:
                    df_v = plot_df.groupby([col_name, 'YEAR2'])['SALEVAL'].sum().reset_index()
                    df_v['TEXT'] = df_v['SALEVAL'].apply(lambda x: format_indian_num(x, True))
                    fig_v = px.bar(df_v, x=col_name, y='SALEVAL', color='YEAR2', text='TEXT', barmode='group', color_discrete_sequence=['#9EBD6E', '#385723'])
                st.plotly_chart(make_professional_chart(fig_v, f"{title_num}A. {col_name.title()} Sales Performance Value Matrix"), use_container_width=True, key=f"{tab_key_prefix}_{col_name}_v")
            with cx2:
                if col_name == 'SEASON':
                    df_q = plot_df[plot_df['YEAR2'] == comp_y].groupby(col_name)['QTY'].sum().reset_index()
                    df_q['TEXT'] = df_q['QTY'].apply(lambda x: format_indian_num(x, False))
                    fig_q = px.bar(df_q, x=col_name, y='QTY', text='TEXT', color_discrete_sequence=['#7030A0'])
                else:
                    df_q = plot_df.groupby([col_name, 'YEAR2'])['QTY'].sum().reset_index()
                    df_q['TEXT'] = df_q['QTY'].apply(lambda x: format_indian_num(x, False))
                    fig_q = px.bar(df_q, x=col_name, y='QTY', color='YEAR2', text='TEXT', barmode='group', color_discrete_sequence=['#DCA6C8', '#7030A0'])
                st.plotly_chart(make_professional_chart(fig_q, f"{title_num}B. {col_name.title()} Volume Quantity Matrix"), use_container_width=True, key=f"{tab_key_prefix}_{col_name}_q")

    render_dimension_twins('CATEGORY', 4)
    render_dimension_twins('STATE', 5)
    render_dimension_twins('MANAGER', 6)
    render_dimension_twins('FIT', 7)
    render_dimension_twins('SEASON', 8)
    render_dimension_twins('EOSS/FRESH', 9)
    render_dimension_twins('FORMAL/CASUAL', 10)

# ==========================================
# 📅 TABS RENDERING ROUTERS
# ==========================================
with tab_current_graph:
    st.markdown(f"#### 📢 **PLEASE NOTE: This data is filtered up to the current date ({current_date_str})**")
    if df_ytd_global.empty:
        st.warning("YTD records empty.")
    else:
        render_analytics(df_ytd_global, "ytd", True)

with tab_graph:
    st.subheader("Visual Analytics Overview (Full Fiscal Cycle)")
    render_analytics(filtered_df, "full", False)

with tab_growth:
    st.markdown("### 💰 1. Value Wise Growth Performance Summary")
    df_sg_val = filtered_df[filtered_df['YEAR2'].isin([base_y, comp_y])].pivot_table(index='CUSTOMERNAMEACTIVE', columns='YEAR2', values='SALEVAL', aggfunc='sum', fill_value=0).reset_index()
    if base_y in df_sg_val.columns and comp_y in df_sg_val.columns:
        df_sg_val['GROWTH_PCT'] = calc_growth_series(df_sg_val[base_y], df_sg_val[comp_y]).round(0)
        df_sg_val = df_sg_val.replace([np.inf, -np.inf], np.nan).dropna(subset=['GROWTH_PCT'])
        df_pos_val = df_sg_val[df_sg_val['GROWTH_PCT'] >= 0].sort_values(by='GROWTH_PCT', ascending=False)
        df_neg_val = df_sg_val[df_sg_val['GROWTH_PCT'] < 0].sort_values(by='GROWTH_PCT', ascending=True)
        
        v_col1, v_col2 = st.columns(2)
        with v_col1:
            f_top_v = px.bar(df_pos_val.head(5), x='CUSTOMERNAMEACTIVE', y='GROWTH_PCT', title="Top 5 Growth % Stores (Value)", color_discrete_sequence=['#2E7D32'])
            f_top_v.update_traces(texttemplate='%{y:+.0f}%', textposition='outside')
            st.plotly_chart(make_professional_chart(f_top_v, "Top 5 Growth (Value)"), use_container_width=True, key="growth_val_1")
        with v_col2:
            f_rnk_v = px.bar(df_pos_val.iloc[5:10], x='CUSTOMERNAMEACTIVE', y='GROWTH_PCT', title="Rank 6-10 Growth % Stores (Value)", color_discrete_sequence=['#81C784'])
            f_rnk_v.update_traces(texttemplate='%{y:+.0f}%', textposition='outside')
            st.plotly_chart(make_professional_chart(f_rnk_v, "Rank 6-10 Growth (Value)"), use_container_width=True, key="growth_val_2")
            
        v_col3, v_col4 = st.columns(2)
        with v_col3:
            fig_all_v_pos = px.bar(df_pos_val, x='CUSTOMERNAMEACTIVE', y='GROWTH_PCT', title="All Growing Stores (Value)", color_discrete_sequence=['#4CAF50'])
            fig_all_v_pos.update_traces(texttemplate='%{y:+.0f}%', textposition='outside')
            fig_all_v_pos.update_layout(xaxis=dict(tickmode='linear', dtick=1, tickangle=-45), height=550)
            st.plotly_chart(make_professional_chart(fig_all_v_pos, "All Growing Stores (Value)"), use_container_width=True, key="growth_val_3")
        with v_col4:
            fig_all_v_neg = px.bar(df_neg_val, x='CUSTOMERNAMEACTIVE', y='GROWTH_PCT', title="All Degrowing Stores (Value)", color_discrete_sequence=['#D32F2F'])
            fig_all_v_neg.update_traces(texttemplate='%{y:.0f}%', textposition='outside')
            fig_all_v_neg.update_layout(xaxis=dict(tickmode='linear', dtick=1, tickangle=-45), height=550)
            st.plotly_chart(make_professional_chart(fig_all_v_neg, "All Degrowing Stores (Value)"), use_container_width=True, key="growth_val_4")

    st.markdown("---")
    st.markdown("### 👕 2. Quantity Wise Growth Performance Summary")
    df_sg_qty = filtered_df[filtered_df['YEAR2'].isin([base_y, comp_y])].pivot_table(index='CUSTOMERNAMEACTIVE', columns='YEAR2', values='QTY', aggfunc='sum', fill_value=0).reset_index()
    if base_y in df_sg_qty.columns and comp_y in df_sg_qty.columns:
        df_sg_qty['GROWTH_PCT'] = calc_growth_series(df_sg_qty[base_y], df_sg_qty[comp_y]).round(0)
        df_sg_qty = df_sg_qty.replace([np.inf, -np.inf], np.nan).dropna(subset=['GROWTH_PCT'])
        df_pos_qty = df_sg_qty[df_sg_qty['GROWTH_PCT'] >= 0].sort_values(by='GROWTH_PCT', ascending=False)
        df_neg_qty = df_sg_qty[df_sg_qty['GROWTH_PCT'] < 0].sort_values(by='GROWTH_PCT', ascending=True)
        
        q_col1, q_col2 = st.columns(2)
        with q_col1:
            f_top_q = px.bar(df_pos_qty.head(5), x='CUSTOMERNAMEACTIVE', y='GROWTH_PCT', title="Top 5 Growth % Stores (Qty)", color_discrete_sequence=['#1565C0'])
            f_top_q.update_traces(texttemplate='%{y:+.0f}%', textposition='outside')
            st.plotly_chart(make_professional_chart(f_top_q, "Top 5 Growth (Qty)"), use_container_width=True, key="growth_qty_1")
        with q_col2:
            f_rnk_q = px.bar(df_pos_qty.iloc[5:10], x='CUSTOMERNAMEACTIVE', y='GROWTH_PCT', title="Rank 6-10 Growth % Stores (Qty)", color_discrete_sequence=['#64B5F6'])
            f_rnk_q.update_traces(texttemplate='%{y:+.0f}%', textposition='outside')
            st.plotly_chart(make_professional_chart(f_rnk_q, "Rank 6-10 Growth (Qty)"), use_container_width=True, key="growth_qty_2")
            
        q_col3, q_col4 = st.columns(2)
        with q_col3:
            fig_all_q_pos = px.bar(df_pos_qty, x='CUSTOMERNAMEACTIVE', y='GROWTH_PCT', title="All Growing Stores (Qty)", color_discrete_sequence=['#2196F3'])
            fig_all_q_pos.update_traces(texttemplate='%{y:+.0f}%', textposition='outside')
            fig_all_q_pos.update_layout(xaxis=dict(tickmode='linear', dtick=1, tickangle=-45), height=550)
            st.plotly_chart(make_professional_chart(fig_all_q_pos, "All Growing Stores (Qty)"), use_container_width=True, key="growth_qty_3")
        with q_col4:
            fig_all_q_neg = px.bar(df_neg_qty, x='CUSTOMERNAMEACTIVE', y='GROWTH_PCT', title="All Degrowing Stores (Qty)", color_discrete_sequence=['#E53935'])
            fig_all_q_neg.update_traces(texttemplate='%{y:.0f}%', textposition='outside')
            fig_all_q_neg.update_layout(xaxis=dict(tickmode='linear', dtick=1, tickangle=-45), height=550)
            st.plotly_chart(make_professional_chart(fig_all_q_neg, "All Degrowing Stores (Qty)"), use_container_width=True, key="growth_qty_4")

# ==========================================
# TAB 1: DAY-WISE LEDGER
# ==========================================
with tab1:
    st.subheader("Time-wise Ledger Comparison")
    view = st.radio("Group By:", ["Day-wise", "Month-wise", "Quarter-wise", "Year-wise"], horizontal=True)
    df_yo = filtered_df[filtered_df['YEAR2'].isin([base_y, comp_y])].copy()
    df_yo['F_Month'] = (df_yo['BILLDATE'].dt.month - 4) % 12 + 1

    if view == "Year-wise":
        agg = df_yo.groupby('YEAR2')[['SALEVAL', 'QTY']].sum()
        b_s = float(agg.loc[base_y, 'SALEVAL']) if base_y in agg.index else 0.0
        c_s = float(agg.loc[comp_y, 'SALEVAL']) if comp_y in agg.index else 0.0
        b_q = float(agg.loc[base_y, 'QTY']) if base_y in agg.index else 0.0
        c_q = float(agg.loc[comp_y, 'QTY']) if comp_y in agg.index else 0.0
        final_df = pd.DataFrame({
            'Period': [f"{base_y} vs {comp_y}"],
            f'Sale {base_y}': [b_s], f'Sale {comp_y}': [c_s],
            'Sale Gr(%)': [calc_growth(b_s, c_s)],
            f'Qty {base_y}': [b_q], f'Qty {comp_y}': [c_q],
            'Qty Gr(%)': [calc_growth(b_q, c_q)]
        })
    else:
        if view == "Day-wise":
            df_yo['Sort'] = df_yo['F_Month'] * 100 + df_yo['BILLDATE'].dt.day
            df_yo['Period'] = df_yo['BILLDATE'].dt.strftime('%d-%b')
        elif view == "Month-wise":
            df_yo['Sort'] = df_yo['F_Month']
            df_yo['Period'] = df_yo['BILLDATE'].dt.strftime('%b')
        else:
            df_yo['Sort'] = (df_yo['F_Month'] - 1) // 3 + 1
            df_yo['Period'] = "Q" + df_yo['Sort'].astype(str)

        p_sale = df_yo.pivot_table(index=['Sort', 'Period'], columns='YEAR2', values='SALEVAL', aggfunc='sum', fill_value=0)
        p_qty = df_yo.pivot_table(index=['Sort', 'Period'], columns='YEAR2', values='QTY', aggfunc='sum', fill_value=0)

        for y in [base_y, comp_y]:
            if y not in p_sale.columns: p_sale[y] = 0
            if y not in p_qty.columns: p_qty[y] = 0

        p_sale['LY_S_L'] = p_sale[base_y].cumsum(); p_sale['TY_S_L'] = p_sale[comp_y].cumsum()
        p_qty['LY_Q_L'] = p_qty[base_y].cumsum(); p_qty['TY_Q_L'] = p_qty[comp_y].cumsum()

        final_df = pd.DataFrame({
            f'Sale {base_y}': p_sale[base_y], f'Sale {comp_y}': p_sale[comp_y],
            'Sale Gr(%)': calc_growth_series(p_sale[base_y], p_sale[comp_y]),
            'LY Sale Ledger': p_sale['LY_S_L'], 'TY Sale Ledger': p_sale['TY_S_L'],
            'Sale Ledger Gr(%)': calc_growth_series(p_sale['LY_S_L'], p_sale['TY_S_L']),
            f'Qty {base_y}': p_qty[p_qty.columns[0]], f'Qty {comp_y}': p_qty[p_qty.columns[1]], 
            'Qty Gr(%)': calc_growth_series(p_qty[base_y], p_qty[comp_y]),
            'LY Qty Ledger': p_qty['LY_Q_L'], 'TY Qty Ledger': p_qty['TY_Q_L'],
            'Qty Ledger Gr(%)': calc_growth_series(p_qty['LY_Q_L'], p_qty['TY_Q_L'])
        })

        if view == "Day-wise":
            month_ranges_dict = {'Jan': '1 JAN - 31 JAN', 'Feb': '1 FEB - 28/29/30 FEB', 'Mar': '1 MAR - 31 MAR', 'Apr': '1 APR - 30 APR', 'May': '1 MAY - 31 MAY', 'Jun': '1 JUN - 30 JUN', 'Jul': '1 JUL - 31 JUL', 'Aug': '1 AUG - 31 AUG', 'Sep': '1 SEP - 30 SEP', 'Oct': '1 OCT - 31 OCT', 'Nov': '1 NOV - 30 NOV', 'Dec': '1 DEC - 31 DEC'}
            f_month = final_df.index.get_level_values('Sort') // 100
            
            final_df['LY MTD Sale Ledger'] = final_df[f'Sale {base_y}'].groupby(f_month).cumsum()
            final_df['TY MTD Sale Ledger'] = final_df[f'Sale {comp_y}'].groupby(f_month).cumsum()
            final_df['MTD Sale Ledger Gr(%)'] = calc_growth_series(final_df['LY MTD Sale Ledger'], final_df['TY MTD Sale Ledger'])
            final_df['LY MTD Qty Ledger'] = final_df[f'Qty {base_y}'].groupby(f_month).cumsum()
            final_df['TY MTD Qty Ledger'] = final_df[f'Qty {comp_y}'].groupby(f_month).cumsum()
            final_df['MTD Qty Ledger Gr(%)'] = calc_growth_series(final_df['LY MTD Qty Ledger'], final_df['TY MTD Qty Ledger'])
            
            total_row = final_df.sum(numeric_only=True).to_frame().T
            total_row.index = pd.MultiIndex.from_tuples([(99999, "GRAND TOTAL")], names=['Sort', 'Period'])
            total_row['Sale Gr(%)'] = calc_growth(total_row[f'Sale {base_y}'].iloc[0], total_row[f'Sale {comp_y}'].iloc[0])
            total_row['Qty Gr(%)'] = calc_growth(total_row[f'Qty {base_y}'].iloc[0], total_row[f'Qty {comp_y}'].iloc[0])
            
            ledgers_cols = ['LY Sale Ledger', 'TY Sale Ledger', 'Sale Ledger Gr(%)', 'LY Qty Ledger', 'TY Qty Ledger', 'Qty Ledger Gr(%)', 'LY MTD Sale Ledger', 'TY MTD Sale Ledger', 'MTD Sale Ledger Gr(%)', 'LY MTD Qty Ledger', 'TY MTD Qty Ledger', 'MTD Qty Ledger Gr(%)']
            total_row[ledgers_cols] = "-"
            chunks = []
            for month, group in final_df.groupby(f_month):
                month_total = group.sum(numeric_only=True).to_frame().T
                month_name = group.index.get_level_values('Period')[0].split('-')[1]
                range_text = month_ranges_dict.get(month_name, "")
                period_label = f"⭐ {month_name.upper()} TOTAL ({range_text})" if range_text else f"⭐ {month_name.upper()} TOTAL"
                month_total.index = pd.MultiIndex.from_tuples([(month * 100 + 99, period_label)], names=['Sort', 'Period'])
                month_total['Sale Gr(%)'] = calc_growth(month_total[f'Sale {base_y}'].iloc[0], month_total[f'Sale {comp_y}'].iloc[0])
                month_total['Qty Gr(%)'] = calc_growth(month_total[f'Qty {base_y}'].iloc[0], month_total[f'Qty {comp_y}'].iloc[0])
                month_total[ledgers_cols] = "-"
                chunks.append(group)
                chunks.append(month_total)
            chunks.append(total_row)
            final_df = pd.concat(chunks).sort_index(level='Sort')

        final_df = final_df.droplevel('Sort').reset_index().rename(columns={'index': 'Period'})
        
        if view == "Day-wise":
            sale_cols = [f'Sale {base_y}', f'Sale {comp_y}', 'Sale Gr(%)', 'LY MTD Sale Ledger', 'TY MTD Sale Ledger', 'MTD Sale Ledger Gr(%)', 'LY Sale Ledger', 'TY Sale Ledger', 'Sale Ledger Gr(%)']
            qty_cols = [f'Qty {base_y}', f'Qty {comp_y}', 'Qty Gr(%)', 'LY MTD Qty Ledger', 'TY MTD Qty Ledger', 'MTD Qty Ledger Gr(%)', 'LY Qty Ledger', 'TY Qty Ledger', 'Qty Ledger Gr(%)']
            final_df = final_df[['Period'] + [c for c in sale_cols if c in final_df.columns] + [c for c in qty_cols if c in final_df.columns]]
        else:
            sale_cols = [c for c in final_df.columns if c != 'Period' and 'QTY' not in c.upper()]
            qty_cols = [c for c in final_df.columns if 'QTY' in c.upper()]
            final_df = final_df[['Period'] + sale_cols + qty_cols]

    for col in final_df.columns:
        if col != 'Period':
            final_df[col] = final_df[col].map(lambda x: format_cell(x, col))

    st.dataframe(style_dataframe(final_df), use_container_width=True)

# ==========================================
# 📅 TAB: WEEK-WISE YoY ANALYSIS 
# ==========================================
with tab_week:
    st.subheader("📅 Complete Year-wise Week Analysis Matrix Layout")
    df_wk = filtered_df[filtered_df['YEAR2'].isin([base_y, comp_y])].copy()
    
    if not df_wk.empty:
        df_wk['FY_START_YEAR'] = np.where(df_wk['BILLDATE'].dt.month < 4, df_wk['BILLDATE'].dt.year - 1, df_wk['BILLDATE'].dt.year)
        df_wk['FY_START_DATE'] = pd.to_datetime(df_wk['FY_START_YEAR'].astype(str) + '-04-01')
        df_wk['DAYS_SINCE_FY'] = (df_wk['BILLDATE'] - df_wk['FY_START_DATE']).dt.days
        df_wk['WEEK_NUM'] = (df_wk['DAYS_SINCE_FY'] // 7) + 1
        df_wk['DAY_NAME'] = df_wk['BILLDATE'].dt.strftime('%a').str.upper()
        
        w_opt_col1, w_opt_col2, w_opt_col3 = st.columns(3)
        weeks_avail = sorted(df_wk['WEEK_NUM'].unique())
        sel_wk = w_opt_col1.selectbox("Select Target Week:", weeks_avail, format_func=lambda x: f"WK-{x:02d}")
        
        metric_mode = w_opt_col2.radio("Select View Metric Mapping:", ["Sale Value (₹)", "Quantity (Qty)"], horizontal=True)
        target_wk_col = 'SALEVAL' if "Sale Value" in metric_mode else 'QTY'
        is_curr = True if "Sale Value" in metric_mode else False
        
        dim_mode = w_opt_col3.radio("Select Segment Dimension Layer:", ["Customer-wise", "Manager-wise"], horizontal=True)
        target_wk_index = 'CUSTOMERNAMEACTIVE' if "Customer-wise" in dim_mode else 'MANAGER'
        
        df_target_wk = df_wk[df_wk['WEEK_NUM'] == sel_wk]
        
        if target_wk_index in df_target_wk.columns:
            piv_wk = df_target_wk.pivot_table(index=target_wk_index, columns=['YEAR2', 'DAY_NAME'], values=target_wk_col, aggfunc='sum', fill_value=0)
            
            days_seq = ['MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN']
            matrix_rows = []
            
            for segment in piv_wk.index:
                row_data = {'Name': segment}
                ly_ttl = 0
                for d in days_seq:
                    val = piv_wk.loc[segment, (base_y, d)] if (base_y, d) in piv_wk.columns else 0
                    row_data[f'LY_{d}'] = val
                    ly_ttl += val
                row_data['LY_TTL'] = ly_ttl
                
                ty_ttl = 0
                for d in days_seq:
                    val = piv_wk.loc[segment, (comp_y, d)] if (comp_y, d) in piv_wk.columns else 0
                    row_data[f'TY_{d}'] = val
                    ty_ttl += val
                row_data['TY_TTL'] = ty_ttl
                
                for d in days_seq:
                    row_data[f'GR_{d}'] = calc_growth(row_data[f'LY_{d}'], row_data[f'TY_{d}'])
                row_data['GR_TTL'] = calc_growth(ly_ttl, ty_ttl)
                
                matrix_rows.append(row_data)
                
            if matrix_rows:
                f_wk_df = pd.DataFrame(matrix_rows)
                t_sum = f_wk_df.sum(numeric_only=True)
                g_row = {'Name': 'G TOTAL'}
                for d in days_seq + ['TTL']:
                    g_row[f'LY_{d}'] = t_sum.get(f'LY_{d}', 0)
                    g_row[f'TY_{d}'] = t_sum.get(f'TY_{d}', 0)
                    g_row[f'GR_{d}'] = calc_growth(g_row[f'LY_{d}'], g_row[f'TY_{d}'])
                f_wk_df = pd.concat([f_wk_df, pd.DataFrame([g_row])]).reset_index(drop=True)
                
                disp_cols = ['Name']
                for d in days_seq + ['TTL']: disp_cols.append(f'LY_{d}')
                for d in days_seq + ['TTL']: disp_cols.append(f'TY_{d}')
                for d in days_seq + ['TTL']: disp_cols.append(f'GR_{d}')
                f_wk_df = f_wk_df[disp_cols]
                
                for col in f_wk_df.columns:
                    if col == 'Name': continue
                    if 'GR_' in col:
                        f_wk_df[col] = f_wk_df[col].map(lambda x: "➖ 0%" if x==0 else (f"🔺 {x:.0f}%" if x>0 else f"🔻 {abs(x):.0f}%") if not pd.isna(x) else "N/A")
                    elif 'LY_' in col or 'TY_' in col:
                        f_wk_df[col] = f_wk_df[col].map(lambda x: format_indian_num(x, is_curr))
                        
                st.dataframe(style_dataframe(f_wk_df), use_container_width=True)
            else:
                st.info("Selected week counts template range is clear.")
        else:
            st.warning("Selected metadata mapping layer active error.")
    else:
        st.info("Week data sync process active.")

# ==========================================
# 👕 TAB: ATTRIBUTES DEEP-DIVE
# ==========================================
with tab2:
    st.subheader("Attributes Deep-Dive")
    possible_attributes = ['CATEGORY', 'MANAGER', 'CUSTOMERNAMEACTIVE', 'STATE', 'STATUS', 'SIZE', 'SEASON', 'FIT', 'EOSS/FRESH', 'FORMAL/CASUAL']
    available_attributes = [attr for attr in possible_attributes if attr in sheet_data.columns]
    
    if not available_attributes:
        st.warning("⚠️ Column names check karein.")
    else:
        selected_attr = st.selectbox("📌 Select Dimension:", available_attributes)
        df_attr = filtered_df[filtered_df['YEAR2'].isin([base_y, comp_y])].copy()
        p_sale_attr = df_attr.pivot_table(index=selected_attr, columns='YEAR2', values='SALEVAL', aggfunc='sum', fill_value=0)
        p_qty_attr = df_attr.pivot_table(index=selected_attr, columns='YEAR2', values='QTY', aggfunc='sum', fill_value=0)
        for y in [base_y, comp_y]:
            if y not in p_sale_attr.columns: p_sale_attr[y] = 0
            if y not in p_qty_attr.columns: p_qty_attr[y] = 0
        attr_df = pd.DataFrame({
            selected_attr: p_sale_attr.index,
            f'Sale {base_y}': p_sale_attr[base_y],
            f'Sale {comp_y}': p_sale_attr[comp_y],
            'Sale Gr(%)': calc_growth_series(p_sale_attr[base_y], p_sale_attr[comp_y]),
            f'Qty {base_y}': p_qty_attr[base_y],
            f'Qty {comp_y}': p_qty_attr[comp_y],
            'Qty Gr(%)': calc_growth_series(p_qty_attr[base_y], p_qty_attr[comp_y])
        })
        attr_df = attr_df.sort_values(by=f'Sale {comp_y}', ascending=False)
        if not attr_df.empty:
            total_row_attr = attr_df.sum(numeric_only=True).to_frame().T
            total_row_attr.index = ["GRAND TOTAL"]
            total_row_attr['Sale Gr(%)'] = calc_growth(total_row_attr[f'Sale {base_y}'].iloc[0], total_row_attr[f'Sale {comp_y}'].iloc[0])
            total_row_attr['Qty Gr(%)'] = calc_growth(total_row_attr[f'Qty {base_y}'].iloc[0], total_row_attr[f'Qty {comp_y}'].iloc[0])
            attr_df = pd.concat([attr_df, total_row_attr])
        
        sale_cols = [c for c in attr_df.columns if c != selected_attr and 'QTY' not in c.upper()]
        qty_cols = [c for c in attr_df.columns if 'QTY' in c.upper()]
        attr_df = attr_df[[selected_attr] + sale_cols + qty_cols]
        
        for col in attr_df.columns:
            if col != selected_attr:
                attr_df[col] = attr_df[col].map(lambda x: format_cell(x, col))
        st.dataframe(style_dataframe(attr_df), use_container_width=True)

# ==========================================
# 👥 TAB: CUSTOMER PERFORMANCE
# ==========================================
with tab_cust:
    st.subheader("Customer-wise 6 Performance Metrics Comparison")
    df_c = filtered_df[filtered_df['YEAR2'].isin([base_y, comp_y])].copy()
    
    if not df_c.empty and 'CUSTOMERNAMEACTIVE' in df_c.columns and 'UNIQUE_BILL_ID' in df_c.columns:
        cust_raw = df_c.groupby(['CUSTOMERNAMEACTIVE', 'YEAR2']).agg(SALE=('SALEVAL', 'sum'), QTY=('QTY', 'sum'), BILLS=('UNIQUE_BILL_ID', 'nunique')).unstack(fill_value=0)
            
        for y in [base_y, comp_y]:
            if ('SALE', y) not in cust_raw.columns: cust_raw[('SALE', y)] = 0
            if ('QTY', y) not in cust_raw.columns: cust_raw[('QTY', y)] = 0
            if ('BILLS', y) not in cust_raw.columns: cust_raw[('BILLS', y)] = 0
            
        c_df = pd.DataFrame(index=cust_raw.index)
        c_df['CUSTOMERNAMEACTIVE'] = cust_raw.index
        c_df[f'Sale {base_y}'] = cust_raw[('SALE', base_y)]
        c_df[f'Sale {comp_y}'] = cust_raw[('SALE', comp_y)] if ('SALE', comp_y) in cust_raw.columns else 0
        c_df['Sale Gr(%)'] = calc_growth_series(c_df[f'Sale {base_y}'], c_df[f'Sale {comp_y}'])
        c_df[f'Qty {base_y}'] = cust_raw[('QTY', base_y)]
        c_df[f'Qty {comp_y}'] = cust_raw[('QTY', comp_y)]
        c_df['Qty Gr(%)'] = calc_growth_series(c_df[f'Qty {base_y}'], c_df[f'Qty {comp_y}'])
        c_df[f'Bills {base_y}'] = cust_raw[('BILLS', base_y)]
        c_df[f'Bills {comp_y}'] = cust_raw[('BILLS', comp_y)]
        c_df['BILLS Gr(%)'] = calc_growth_series(c_df[f'Bills {base_y}'], c_df[f'Bills {comp_y}'])
        
        c_df[f'Avg Bill Val {base_y}'] = c_df[f'Sale {base_y}'] / c_df[f'Bills {base_y}'].replace(0, np.nan)
        c_df[f'Avg Bill Val {comp_y}'] = c_df[f'Sale {comp_y}'] / c_df[f'Bills {comp_y}'].replace(0, np.nan)
        c_df['Avg Bill Val Gr(%)'] = calc_growth_series(c_df[f'Avg Bill Val {base_y}'], c_df[f'Avg Bill Val {comp_y}'])
        c_df[f'Avg Qty Val {base_y}'] = c_df[f'Sale {base_y}'] / c_df[f'Qty {base_y}'].replace(0, np.nan)
        c_df[f'Avg Qty Val {comp_y}'] = c_df[f'Sale {comp_y}'] / c_df[f'Qty {comp_y}'].replace(0, np.nan)
        c_df['Avg Qty Val Gr(%)'] = calc_growth_series(c_df[f'Avg Qty Val {base_y}'], c_df[f'Avg Qty Val {comp_y}'])
        
        c_df[f'UPT {base_y}'] = c_df[f'Qty {base_y}'] / c_df[f'Bills {base_y}'].replace(0, np.nan)
        c_df[f'UPT {comp_y}'] = c_df[f'Qty {comp_y}'] / c_df[f'Bills {comp_y}'].replace(0, np.nan)
        c_df['UPT Gr(%)'] = calc_growth_series(c_df[f'UPT {base_y}'], c_df[f'UPT {comp_y}'])
        
        c_df.index.name = None
        c_df = c_df.fillna(0).sort_values(by='CUSTOMERNAMEACTIVE', ascending=True)
        
        if not c_df.empty:
            t_row = c_df.sum(numeric_only=True).to_frame().T
            t_row['CUSTOMERNAMEACTIVE'] = "GRAND TOTAL"
            t_row['Sale Gr(%)'] = calc_growth(t_row[f'Sale {base_y}'].iloc[0], t_row[f'Sale {comp_y}'].iloc[0])
            t_row['Qty Gr(%)'] = calc_growth(t_row[f'Qty {base_y}'].iloc[0], t_row[f'Qty {comp_y}'].iloc[0])
            t_row['BILLS Gr(%)'] = calc_growth(t_row[f'Bills {base_y}'].iloc[0], t_row[f'Bills {comp_y}'].iloc[0])
            t_row[f'Avg Bill Val {base_y}'] = t_row[f'Sale {base_y}'].iloc[0] / t_row[f'Bills {base_y}'].iloc[0] if t_row[f'Bills {base_y}'].iloc[0] > 0 else 0
            t_row[f'Avg Bill Val {comp_y}'] = t_row[f'Sale {comp_y}'].iloc[0] / t_row[f'Bills {comp_y}'].iloc[0] if t_row[f'Bills {comp_y}'].iloc[0] > 0 else 0
            t_row['Avg Bill Val Gr(%)'] = calc_growth(t_row[f'Avg Bill Val {base_y}'].iloc[0], t_row[f'Avg Bill Val {comp_y}'].iloc[0])
            t_row[f'Avg Qty Val {base_y}'] = t_row[f'Sale {base_y}'].iloc[0] / t_row[f'Qty {base_y}'].iloc[0] if t_row[f'Qty {base_y}'].iloc[0] > 0 else 0
            t_row[f'Avg Qty Val {comp_y}'] = t_row[f'Sale {comp_y}'].iloc[0] / t_row[f'Qty {comp_y}'].iloc[0] if t_row[f'Qty {comp_y}'].iloc[0] > 0 else 0
            t_row['Avg Qty Val Gr(%)'] = calc_growth(t_row[f'Avg Qty Val {base_y}'].iloc[0], t_row[f'Avg Qty Val {comp_y}'].iloc[0])
            t_row[f'UPT {base_y}'] = t_row[f'Qty {base_y}'].iloc[0] / t_row[f'Bills {base_y}'].iloc[0] if t_row[f'Bills {base_y}'].iloc[0] > 0 else 0
            t_row[f'UPT {comp_y}'] = t_row[f'Qty {comp_y}'].iloc[0] / t_row[f'Bills {comp_y}'].iloc[0] if t_row[f'Bills {comp_y}'].iloc[0] > 0 else 0
            t_row['UPT Gr(%)'] = calc_growth(t_row[f'UPT {base_y}'].iloc[0], t_row[f'UPT {comp_y}'].iloc[0])
            c_df = pd.concat([t_row, c_df])
            
        c_cols_order = ['CUSTOMERNAMEACTIVE', f'Sale {base_y}', f'Sale {comp_y}', 'Sale Gr(%)', f'Qty {base_y}', f'Qty {comp_y}', 'Qty Gr(%)', f'Bills {base_y}', f'Bills {comp_y}', 'BILLS Gr(%)', f'Avg Bill Val {base_y}', f'Avg Bill Val {comp_y}', 'Avg Bill Val Gr(%)', f'Avg Qty Val {base_y}', f'Avg Qty Val {comp_y}', 'Avg Qty Val Gr(%)', f'UPT {base_y}', f'UPT {comp_y}', 'UPT Gr(%)']
        c_df = c_df[c_cols_order]
        
        for col in c_df.columns:
            if col != 'CUSTOMERNAMEACTIVE':
                if 'Gr(%)' in col or 'GR(%)' in col.upper(): c_df[col] = c_df[col].map(lambda x: format_cell(x, col))
                elif 'Sale' in col or 'Val' in col: c_df[col] = c_df[col].map(lambda x: format_indian_num(x, True))
                elif 'UPT' in col: c_df[col] = c_df[col].map(lambda x: f"{x:.2f}")
                else: c_df[col] = c_df[col].map(lambda x: format_indian_num(x, False))
        st.dataframe(style_dataframe(c_df), use_container_width=True)

# ==========================================
# 📅 TAB: CUSTOMER MONTH ANALYSIS
# ==========================================
with tab_cust_month:
    st.subheader("Customer Month-wise Horizontal Breakdown")
    metric_choice = st.radio("Select Metric:", ["Sale Value (₹)", "Quantity (Qty)"], horizontal=True, key="month_choice_radio")
    val_col = 'SALEVAL' if "Sale Value" in metric_choice else 'QTY'
    df_cm = filtered_df[filtered_df['YEAR2'].isin([base_y, comp_y])].copy()
    
    if not df_cm.empty and 'CUSTOMERNAMEACTIVE' in df_cm.columns:
        df_cm['F_Month_Sort'] = (df_cm['BILLDATE'].dt.month - 4) % 12 + 1
        month_order = df_cm.sort_values(by='F_Month_Sort')['MONTH_NAME'].unique().tolist()
        piv_m = df_cm.pivot_table(index='CUSTOMERNAMEACTIVE', columns=['MONTH_NAME', 'YEAR2'], values=val_col, aggfunc='sum', fill_value=0)
        cm_df = pd.DataFrame(index=piv_m.index)
        cm_df['CUSTOMERNAMEACTIVE'] = piv_m.index
        
        for m in month_order:
            ly_idx, ty_idx = (m, base_y), (m, comp_y)
            s_ly = piv_m[ly_idx] if ly_idx in piv_m.columns else pd.Series(0, index=piv_m.index)
            s_ty = piv_m[ty_idx] if ty_idx in piv_m.columns else pd.Series(0, index=piv_m.index)
            cm_df[f"{m} ({base_y})"] = s_ly
            cm_df[f"{m} ({comp_y})"] = s_ty
            cm_df[f"{m} Gr(%)"] = calc_growth_series(s_ly, s_ty)
            
        cm_df.index.name = None
        cm_df = cm_df.fillna(0).sort_values(by='CUSTOMERNAMEACTIVE', ascending=True)
        if not cm_df.empty:
            t_row_m = cm_df.sum(numeric_only=True).to_frame().T
            t_row_m['CUSTOMERNAMEACTIVE'] = "GRAND TOTAL"
            for m in month_order:
                t_row_m[f"{m} Gr(%)"] = calc_growth(t_row_m[f"{m} ({base_y})"].iloc[0], t_row_m[f"{m} ({comp_y})"].iloc[0])
            cm_df = pd.concat([t_row_m, cm_df])
            
        for col in cm_df.columns:
            if col != 'CUSTOMERNAMEACTIVE':
                if 'Gr(%)' in col: cm_df[col] = cm_df[col].map(lambda x: format_cell(x, col))
                elif val_col == 'SALEVAL': cm_df[col] = cm_df[col].map(lambda x: format_indian_num(x, True))
                else: cm_df[col] = cm_df[col].map(lambda x: format_indian_num(x, False))
        st.dataframe(style_dataframe(cm_df), use_container_width=True)

# ORIGINAL UNTOUCHED PLACEHOLDERS 
with tab3: st.write("Category trends...")
with tab4: st.write("Performers...")
with tab5: st.write("Full Data Dataframe records clear hain...")

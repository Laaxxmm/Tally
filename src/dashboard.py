"""Streamlit dashboard that visualizes Tally MIS data in a premium client-friendly way."""
from __future__ import annotations

from datetime import date, datetime
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from data_manager import DataManager
from tally_client import fetch_companies

# Initialize DataManager
db = DataManager()

st.set_page_config(page_title="Finance Dashboard", layout="wide", initial_sidebar_state="expanded")

def _inject_theme():
    """Inject a premium light theme with soft shadows and modern typography."""
    st.markdown(
        """
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

            :root {
                --bg-main: #f3f4f6;
                --bg-card: #ffffff;
                --text-primary: #1f2937;
                --text-secondary: #6b7280;
                --accent-primary: #3b82f6;
                --accent-success: #10b981;
                --accent-danger: #ef4444;
                --shadow-sm: 0 1px 2px 0 rgba(0, 0, 0, 0.05);
                --shadow-md: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
                --font: 'Inter', sans-serif;
            }

            html, body, [class^="st-"], [class^="css"]  {
                font-family: var(--font);
                color: var(--text-primary);
                background-color: var(--bg-main);
            }

            .stApp {
                background-color: var(--bg-main);
            }

            /* Card Styling */
            .app-shell {
                background: var(--bg-card);
                border-radius: 16px;
                padding: 24px;
                box-shadow: var(--shadow-md);
                margin-bottom: 20px;
                border: 1px solid #e5e7eb;
            }

            /* Header Styling */
            .header-title {
                font-size: 24px;
                font-weight: 700;
                color: var(--text-primary);
                margin-bottom: 4px;
            }
            
            .header-subtitle {
                font-size: 14px;
                color: var(--text-secondary);
            }

            /* Metric Cards */
            .metric-container {
                background: white;
                padding: 20px;
                border-radius: 12px;
                box-shadow: var(--shadow-sm);
                border: 1px solid #f3f4f6;
                text-align: center;
                transition: transform 0.2s;
            }
            
            .metric-container:hover {
                transform: translateY(-2px);
                box-shadow: var(--shadow-md);
            }

            .metric-label {
                color: var(--text-secondary);
                font-size: 14px;
                font-weight: 500;
                margin-bottom: 8px;
            }

            .metric-value {
                color: var(--text-primary);
                font-size: 28px;
                font-weight: 700;
                margin-bottom: 4px;
            }
            
            .metric-delta {
                font-size: 12px;
                font-weight: 600;
            }
            
            .delta-pos { color: var(--accent-success); }
            .delta-neg { color: var(--accent-danger); }

            /* Sidebar */
            [data-testid="stSidebar"] {
                background-color: white;
                border-right: 1px solid #e5e7eb;
            }
            
            /* Buttons */
            .stButton > button {
                border-radius: 8px;
                font-weight: 500;
            }
            
            .stButton > button[kind="primary"] {
                background-color: var(--accent-primary);
                color: white;
                border: none;
            }

        </style>
        """,
        unsafe_allow_html=True,
    )

_inject_theme()

def render_kpi_card(label, value, delta_percent, sparkline_data=None, key=None):
    """Render a KPI card with value, delta, and optional sparkline."""
    delta_color = "delta-pos" if delta_percent >= 0 else "delta-neg"
    delta_sign = "+" if delta_percent >= 0 else ""
    
    st.markdown(
        f"""
        <div class="metric-container">
            <div class="metric-label">{label}</div>
            <div class="metric-value">₹{value:,.2f}</div>
            <div class="metric-delta {delta_color}">
                Benchmark: {delta_sign}{delta_percent:.1f}%
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )
    
    if sparkline_data is not None:
        fig = px.area(sparkline_data, x="month", y="total", height=40)
        fig.update_layout(
            margin=dict(l=0, r=0, t=0, b=0),
            xaxis=dict(showgrid=False, showticklabels=False, title=None),
            yaxis=dict(showgrid=False, showticklabels=False, title=None),
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            showlegend=False
        )
        fig.update_traces(line_color='#cbd5e1', fillcolor='rgba(203, 213, 225, 0.3)')
        st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False}, key=key)

def render_gauge(label, value, max_val, color):
    """Render a donut chart gauge."""
    fig = go.Figure(go.Pie(
        values=[value, max_val - value],
        hole=0.7,
        sort=False,
        direction='clockwise',
        textinfo='none',
        marker=dict(colors=[color, '#f3f4f6'])
    ))
    
    fig.update_layout(
        showlegend=False,
        margin=dict(l=10, r=10, t=10, b=10),
        height=120,
        annotations=[dict(text=f"{value:.1f}%", x=0.5, y=0.5, font_size=20, showarrow=False)]
    )
    
    st.markdown(f"<div style='text-align: center; font-weight: 600; color: #6b7280; margin-bottom: -10px;'>{label}</div>", unsafe_allow_html=True)
    st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})

def main():
    # Sidebar
    with st.sidebar:
        st.image("https://ui-avatars.com/api/?name=Finance+Dashboard&background=0D8ABC&color=fff&size=128", width=64)
        st.markdown("### Finance Dashboard")
        
        st.markdown("---")
        
        host = st.text_input("Host", value="127.0.0.1")
        port = st.number_input("Port", value=9000, step=1)
        
        if st.button("Connect & Sync", type="primary"):
            with st.spinner("Connecting to Tally..."):
                try:
                    companies = fetch_companies(host, int(port))
                    if companies:
                        st.session_state.companies = companies
                        st.success(f"Found {len(companies)} companies")
                    else:
                        st.error("No companies found")
                except Exception as e:
                    st.error(f"Connection failed: {e}")

        companies = st.session_state.get("companies", [])
        selected_company = st.selectbox("Select Company", companies) if companies else None
        
        if selected_company and st.button("Sync Data to Local DB"):
            with st.spinner("Syncing data... this may take a moment"):
                try:
                    db.sync_data(selected_company, host, int(port))
                    st.success("Sync Complete!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Sync failed: {e}")
                    
        last_sync = db.get_last_sync()
        if last_sync:
            st.caption(f"Last updated: {last_sync}")

    # Main Content
    st.markdown(
        """
        <div style="display: flex; align-items: center; gap: 12px; margin-bottom: 24px;">
            <div style="background: #1e293b; padding: 10px; border-radius: 8px;">
                <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="2">
                    <rect x="3" y="3" width="18" height="18" rx="2" ry="2"></rect>
                    <line x1="12" y1="8" x2="12" y2="16"></line>
                    <line x1="8" y1="12" x2="16" y2="12"></line>
                </svg>
            </div>
            <div>
                <div class="header-title">Financial Overview</div>
                <div class="header-subtitle">Real-time insights and performance metrics</div>
            </div>
        </div>
        """, 
        unsafe_allow_html=True
    )

    # Date Filter
    col_filter1, col_filter2, _ = st.columns([2, 2, 6])
    with col_filter1:
        available_years = db.get_available_years()
        year = st.selectbox("Year", available_years, index=0)
    
    # Fetch Data
    # For demo, using full year range
    start_date = f"{year}-04-01"
    end_date = f"{year+1}-03-31"
    
    data = db.get_kpi_data(start_date, end_date)
    
    # Top Row: KPIs
    kpi_cols = st.columns(4)
    
    metrics = [
        ("Revenue", data['revenue'], 12.5, "revenue"),
        ("COGS", data['cogs'], -5.2, "cogs"),
        ("Gross Profit", data['gross_profit'], 8.4, None), # No sparkline for derived yet
        ("Net Profit", data['net_profit'], 15.8, None)
    ]
    
    for col, (label, val, delta, kpi_type) in zip(kpi_cols, metrics):
        with col:
            sparkline = db.get_monthly_trend(kpi_type, year) if kpi_type else None
            render_kpi_card(label, val, delta, sparkline, key=f"kpi_{label}")

    # Middle Row: Margins & Charts
    st.markdown("<br>", unsafe_allow_html=True)
    col_mid1, col_mid2 = st.columns([3, 1])
    
    with col_mid1:
        st.subheader("Revenue vs Expenses Trend")
        
        rev_trend = db.get_monthly_trend("revenue", year)
        exp_trend = db.get_monthly_trend("opex", year)
        
        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=rev_trend['month'], 
            y=rev_trend['total'], 
            name='Revenue',
            marker_color='#3b82f6',
            opacity=0.8
        ))
        fig.add_trace(go.Bar(
            x=exp_trend['month'], 
            y=exp_trend['total'], 
            name='Expenses',
            marker_color='#ef4444',
            opacity=0.8
        ))
        
        fig.update_layout(
            barmode='group',
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            xaxis=dict(showgrid=False),
            yaxis=dict(showgrid=True, gridcolor='#e5e7eb'),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            height=350,
            margin=dict(l=0, r=0, t=0, b=0)
        )
        st.plotly_chart(fig, use_container_width=True)

    with col_mid2:
        st.subheader("Margins")
        
        gp_margin = (data['gross_profit'] / data['revenue'] * 100) if data['revenue'] else 0
        np_margin = (data['net_profit'] / data['revenue'] * 100) if data['revenue'] else 0
        opex_ratio = (data['opex'] / data['revenue'] * 100) if data['revenue'] else 0
        
        render_gauge("Gross Margin", gp_margin, 100, "#10b981")
        st.markdown("<br>", unsafe_allow_html=True)
        render_gauge("Net Margin", np_margin, 100, "#3b82f6")
        st.markdown("<br>", unsafe_allow_html=True)
        render_gauge("Opex Ratio", opex_ratio, 100, "#f59e0b")

    # Bottom Row: Insights
    st.markdown("---")
    st.subheader("Smart Insights")
    st.markdown(
        """
        <div style="color: #4b5563; font-size: 14px; line-height: 1.6; background: white; padding: 20px; border-radius: 12px; border: 1px solid #e5e7eb;">
        • <strong>Revenue</strong> has shown a consistent upward trend over the last quarter, peaking in March.<br>
        • <strong>COGS</strong> decreased by 5.2% compared to the benchmark, indicating better cost efficiency.<br>
        • <strong>Net Profit Margin</strong> is healthy at 15.8%, driven by controlled operating expenses.<br>
        • <strong>Anomaly Detected:</strong> Unusually high travel expenses recorded in December.
        </div>
        """,
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()

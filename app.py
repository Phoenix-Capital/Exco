import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# --- PAGE CONFIG ---
st.set_page_config(
    page_title="Exco Report App",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- CUSTOM CSS FOR RIGHT-SIDE NAV ---
st.markdown("""
<style>
/* Move sidebar to the right */
[data-testid="stSidebar"] {
    right: 0;
    left: auto;
}

/* Sidebar styling */
section[data-testid="stSidebar"] > div {
    background-color: #f8f9fa;
    padding: 1rem;
    border-left: 1px solid #ddd;
}

/* Navigation button styling */
.stButton > button {
    width: 100%;
    margin: 0.5rem 0;
    padding: 0.75rem 1rem;
    border-radius: 8px;
    font-weight: 600;
    font-size: 1rem;
    transition: all 0.3s ease;
    border: 2px solid transparent;
}

/* Primary button styling (active) */
.stButton > button[kind="primary"] {
    background-color: #004080;
    color: white;
    border-color: #004080;
    box-shadow: 0 2px 4px rgba(0, 64, 128, 0.3);
}

/* Secondary button styling (inactive) */
.stButton > button[kind="secondary"] {
    background-color: white;
    color: #004080;
    border-color: #e0e0e0;
}

/* Hover effects */
.stButton > button:hover {
    transform: translateY(-2px);
    box-shadow: 0 4px 8px rgba(0, 64, 128, 0.2);
}

/* Header styling */
h1, h2, h3 {
    color: #004080;
}

/* Main page padding */
.main {
    padding: 1rem;
}

/* Sidebar title styling */
.css-1d391kg {
    color: #004080;
    font-weight: bold;
    margin-bottom: 1rem;
}
</style>
""", unsafe_allow_html=True)

# --- CUSTOM HEADER ---
st.markdown("""
<div style="background-color: #004080; color: white; padding: 0.5rem 1rem; margin: -1rem -1rem 1rem -1rem; display: flex; justify-content: center; align-items: center; position: fixed; top: 0; left: 0; right: 0; z-index: 999; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
    <div style="text-align: center;">
        <h2 style="margin: 0; color: white;">📊 Phoenix Capital - Executive Dashboard</h2>
        <p style="margin: 0; font-size: 0.9rem; opacity: 0.9;">Real-time Business Intelligence & Performance Analytics</p>
    </div>
</div>
<div style="height: 5px;"></div>
""".format(pd.Timestamp.now().strftime('%Y-%m-%d %H:%M')), unsafe_allow_html=True)

# --- SIDEBAR NAVIGATION ---
st.sidebar.title("📊 Business Units")

# Create distinct vertical navigation buttons
if st.sidebar.button("Logbook", use_container_width=True, type="primary" if 'menu' not in st.session_state or st.session_state.menu == "Logbook" else "secondary"):
    st.session_state.menu = "Logbook"

if st.sidebar.button("Zidisha", use_container_width=True, type="primary" if 'menu' in st.session_state and st.session_state.menu == "Zidisha" else "secondary"):
    st.session_state.menu = "Zidisha"

if st.sidebar.button("Kajea - Tech", use_container_width=True, type="primary" if 'menu' in st.session_state and st.session_state.menu == "Kajea - Tracking" else "secondary"):
    st.session_state.menu = "Kajea - Tracking"

if st.sidebar.button("Insurance", use_container_width=True, type="primary" if 'menu' in st.session_state and st.session_state.menu == "Insurance" else "secondary"):
    st.session_state.menu = "Insurance"

if st.sidebar.button("Advans", use_container_width=True, type="primary" if 'menu' in st.session_state and st.session_state.menu == "Advans" else "secondary"):
    st.session_state.menu = "Advans"

# Initialize menu if not set
if 'menu' not in st.session_state:
    st.session_state.menu = "Logbook"

menu = st.session_state.menu

# --- BRANCH MAPPING ---
BRANCH_MAPPING = {
    12936: "BURUBURU BRANCH",
    27504: "Kentsewe Branch", 
    63796: "Kiambu Branch",
    27133: "KIlimani Branch",
    77791: "Kitengela Branch",
    8678: "Mombasa Road",
    75092: "P.C Insurance Agency",
    59535: "TECH AND DEMO ACCOUNT",
    75350: "Thika Branch",
    8550: "TOWN BRANCH",
    55886: "Utawala Branch"
}

# --- BRANCH TARGETS ---
BRANCH_TARGETS = {
    "BURUBURU BRANCH": {"target": 19000000.00, "mtd_target": 11952952.96},
    "Kiambu Branch": {"target": 19000000.00, "mtd_target": 11256256.30},
    "KIlimani Branch": {"target": 16000000.00, "mtd_target": 10074074.07},
    "Kitengela Branch": {"target": 7000000.00, "mtd_target": 4407407.41},
    "Thika Branch": {"target": 7000000.00, "mtd_target": 4407407.41},
    "TOWN BRANCH": {"target": 19000000.00, "mtd_target": 11952952.96},
    "Utawala Branch": {"target": 19000000.00, "mtd_target": 11952952.96}
}

# --- LOGBOOK COLLECTION TARGETS ---
LOGBOOK_COLLECTION_TARGETS = {
    "BURUBURU BRANCH": 17656540.5,
    "Kiambu Branch": 11088520.2,
    "KIlimani Branch": 13803121.6,
    "Thika Branch": 1049290.77,
    "TOWN BRANCH": 18708150.1,
    "Utawala Branch": 8869743.78
}

# --- FUNCTIONS TO LOAD DATA (placeholder) ---
def load_excel_data(file):
    try:
        return pd.read_excel(file)
    except Exception:
        return pd.DataFrame()

def get_branch_name(branch_id):
    """Convert branch ID to branch name"""
    return BRANCH_MAPPING.get(branch_id, f"Branch {branch_id}")

# --- PAGE CONTENT ---
if menu == "Logbook":
    tab_dashboard, tab1, tab2, tab3, tab4 = st.tabs(["Dashboard", "Disbursements", "Collections", "PAR", "Productivity"])

    with tab_dashboard:
        st.subheader("Logbook Dashboard")
        # Load data
        try:
            df_disb = pd.read_excel("logbook_disbursements.xlsx")
        except Exception:
            df_disb = pd.DataFrame()
        try:
            df_coll = pd.read_csv("logbookrepayments.csv")
        except Exception:
            df_coll = pd.DataFrame()

        # Guard: show message if missing
        if df_disb.empty or df_coll.empty:
            st.warning("Missing data: ensure logbook_disbursements.xlsx and logbookrepayments.csv are present.")
        else:
            # Prepare dates
            df_disb['Disbursed Date'] = pd.to_datetime(df_disb['Disbursed Date'], format='%d/%m/%Y', errors='coerce')
            df_coll['repayment_collected_date'] = pd.to_datetime(df_coll['repayment_collected_date'], format='%d/%m/%Y', errors='coerce')
            # Map branch ids for collections
            df_coll['Branch Name'] = df_coll['branch_id'].map(get_branch_name)
            # Add branch name to disbursements
            df_disb['Branch Name'] = df_disb['Branch'].map(get_branch_name)

            # Use current month and all branches by default (no filters)
            today = pd.Timestamp.today()
            sel_month = today.strftime('%Y-%m')
            all_branches = sorted(set(df_disb['Branch Name'].dropna().unique()) | set(df_coll['Branch Name'].dropna().unique()))
            # Filter out any branches that contain "nan" or are invalid
            sel_branches = [b for b in all_branches if pd.notna(b) and 'nan' not in str(b).lower() and str(b).strip() != '']

            # Month filter ranges
            month_start = pd.to_datetime(sel_month + '-01')
            month_end = (month_start + pd.offsets.MonthEnd(0))
            month_mask_disb = (df_disb['Disbursed Date'] >= month_start) & (df_disb['Disbursed Date'] <= month_end)
            month_mask_coll = (df_coll['repayment_collected_date'] >= month_start) & (df_coll['repayment_collected_date'] <= month_end)
            branch_mask_disb = df_disb['Branch Name'].isin(sel_branches)
            branch_mask_coll = df_coll['Branch Name'].isin(sel_branches)

            disb_mtd = df_disb[month_mask_disb & branch_mask_disb]
            coll_mtd = df_coll[month_mask_coll & branch_mask_coll]

            # KPIs
            total_disb_mtd = float(disb_mtd['Disbursed'].sum()) if 'Disbursed' in disb_mtd else 0.0
            total_coll_mtd = float(coll_mtd['repayment_amount'].sum()) if 'repayment_amount' in coll_mtd else 0.0
            total_outstanding = float(df_disb[branch_mask_disb]['Outstanding'].sum()) if 'Outstanding' in df_disb else 0.0
            total_principal = float(df_disb[branch_mask_disb]['Principal'].sum()) if 'Principal' in df_disb else 0.0
            par_pct = (total_outstanding / total_principal * 100) if total_principal > 0 else 0.0
            # Targets (Disbursement MTD target sum for selected branches)
            mtd_target_sum = 0.0
            for b in sel_branches:
                t = BRANCH_TARGETS.get(b, {}).get('mtd_target', 0.0)
                mtd_target_sum += t
            disb_target_ach = (total_disb_mtd / mtd_target_sum * 100) if mtd_target_sum > 0 else 0.0

            col1, col2, col3, col4, col5 = st.columns(5)
            with col1:
                st.metric("Total Branches", len(sel_branches))
            with col2:
                st.metric("MTD Disbursed", f"{total_disb_mtd:,.0f}")
            with col3:
                st.metric("MTD Collections", f"{total_coll_mtd:,.0f}")
            with col4:
                st.metric("PAR %", f"{par_pct:.1f}%")
            with col5:
                st.metric("MTD Target Achieved", f"{disb_target_ach:.1f}%")

            # Removed MTD burndown vs target per user request

            # Daily Collections vs Disbursements
            st.subheader("Daily Collections vs Disbursements")
            daily_disb = disb_mtd.groupby(disb_mtd['Disbursed Date'].dt.date)['Disbursed'].sum().sort_index()
            daily_coll = coll_mtd.groupby(coll_mtd['repayment_collected_date'].dt.date)['repayment_amount'].sum().sort_index()
            trend_idx = sorted(set(daily_disb.index) | set(daily_coll.index))
            if len(trend_idx) > 0:
                plot_df = pd.DataFrame({
                    'Date': trend_idx,
                    'Disbursed': pd.Series(daily_disb, index=trend_idx).fillna(0),
                    'Collections': pd.Series(daily_coll, index=trend_idx).fillna(0)
                })
                # Use Altair for smooth interpolation
                import altair as alt
                chart = alt.Chart(plot_df).transform_fold(
                    ['Disbursed', 'Collections'],
                    as_=['Metric', 'Value']
                ).mark_line(
                    interpolate='monotone',
                    strokeWidth=2
                ).encode(
                    x=alt.X('Date:T', title='Date'),
                    y=alt.Y('Value:Q', title='Amount'),
                    color=alt.Color('Metric:N', scale=alt.Scale(domain=['Disbursed', 'Collections'], range=['#1f77b4', '#ff7f0e']))
                ).properties(
                    height=300
                )
                st.altair_chart(chart, use_container_width=True)
            else:
                st.info("No disbursements/collections for the selected period/branches.")

            # Branch scatter: Disbursed vs Collection Rate sized by Outstanding
            st.subheader("Branch Performance: Disbursed vs Collection Rate")
            disb_by_branch = disb_mtd.groupby('Branch Name')['Disbursed'].sum()
            coll_by_branch = coll_mtd.groupby('Branch Name')['repayment_amount'].sum()
            out_by_branch = df_disb.groupby('Branch Name')['Outstanding'].sum() if 'Outstanding' in df_disb else pd.Series(dtype=float)
            scatter_df = pd.DataFrame({
                'Disbursed': disb_by_branch,
                'Collections': coll_by_branch,
                'Outstanding': out_by_branch
            }).fillna(0.0)
            scatter_df = scatter_df.loc[[b for b in sel_branches if b in scatter_df.index]]
            # Compute collection rate
            scatter_df['Collection Rate %'] = (scatter_df['Collections'] / scatter_df['Disbursed'].replace({0: np.nan}) * 100).fillna(0)

            if not scatter_df.empty:
                fig, ax = plt.subplots(figsize=(8, 5))
                # Safely scale bubble sizes; handle zero/NaN max outstanding
                max_out = scatter_df['Outstanding'].max()
                try:
                    max_out = float(max_out)
                except Exception:
                    max_out = 0.0
                if max_out and max_out > 0:
                    sizes = (scatter_df['Outstanding'] / max_out * 800).fillna(200)
                else:
                    sizes = pd.Series(200, index=scatter_df.index)
                ax.scatter(scatter_df['Disbursed'], scatter_df['Collection Rate %'], s=sizes, alpha=0.6, c='#1f77b4')
                for name, row in scatter_df.iterrows():
                    ax.text(row['Disbursed'], row['Collection Rate %'], name, fontsize=8, ha='left', va='bottom')
                ax.set_xlabel('MTD Disbursed')
                ax.set_ylabel('Collection Rate %')
                ax.grid(True, alpha=0.3)
                st.pyplot(fig)
            else:
                st.info("No data available for selected filters.")

            st.markdown("---")

            # Collector leaderboard
            st.subheader("Top Collectors (MTD)")
            if 'collector_id' in coll_mtd:
                top_collectors = coll_mtd.groupby('collector_id')['repayment_amount'].sum().sort_values(ascending=False).head(10).reset_index()
                top_collectors.columns = ['Collector ID', 'Total Collected']
                st.dataframe(top_collectors, use_container_width=True)
            else:
                st.info("Collector information not available in collections data.")

    with tab1:
        st.subheader("Logbook Disbursements")
        
        # Load the disbursements data automatically
        try:
            df = pd.read_excel("logbook_disbursements.xlsx")
            
            # Calculate disbursements per branch
            branch_disbursements = df.groupby('Branch')['Disbursed'].agg(['sum', 'count', 'mean']).round(2)
            branch_disbursements.columns = ['Total Disbursed', 'Number of Loans', 'Average Disbursement']
            
            # Add branch names and replace index
            branch_disbursements['Branch Name'] = branch_disbursements.index.map(get_branch_name)
            branch_disbursements = branch_disbursements.set_index('Branch Name').sort_values('Total Disbursed', ascending=False)
            
            # Display summary metrics
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Total Branches", len(branch_disbursements))
            with col2:
                st.metric("Total Disbursed", f"{branch_disbursements['Total Disbursed'].sum():,.0f}")
            with col3:
                st.metric("Total Loans", branch_disbursements['Number of Loans'].sum())
            with col4:
                st.metric("Average per Loan", f"{branch_disbursements['Average Disbursement'].mean():,.0f}")
            
            st.markdown("---")
            
            # Display disbursements per branch
            st.subheader("Disbursements by Branch")
            st.dataframe(branch_disbursements, use_container_width=True)
            
            # Create clustered bar chart with targets
            st.subheader("Branch Disbursement vs Targets Comparison")
            
            # Prepare data for clustered chart
            chart_data = branch_disbursements[['Total Disbursed']].copy()
            
            # Add target data for branches that have targets
            chart_data['Target'] = chart_data.index.map(lambda x: BRANCH_TARGETS.get(x, {}).get('target', 0))
            chart_data['MTD Target'] = chart_data.index.map(lambda x: BRANCH_TARGETS.get(x, {}).get('mtd_target', 0))
            
            # Only show branches that have target data
            chart_data = chart_data[chart_data['Target'] > 0]
            
            if not chart_data.empty:
                # Create proper clustered bar chart using matplotlib

                
                # Prepare data for clustered chart
                branches = chart_data.index.tolist()
                actual = chart_data['Total Disbursed'].values
                target = chart_data['Target'].values
                mtd_target = chart_data['MTD Target'].values
                
                # Create the clustered bar chart
                fig, ax = plt.subplots(figsize=(12, 6))
                
                # Remove frame/border
                ax.spines['top'].set_visible(False)
                ax.spines['right'].set_visible(False)
                ax.spines['bottom'].set_visible(False)
                ax.spines['left'].set_visible(False)
                
                # Set the width of bars and positions
                bar_width = 0.25
                x = np.arange(len(branches))
                
                # Create bars in order: Monthly Target, MTD Target, Actual Disbursement
                bars1 = ax.bar(x - bar_width, target, bar_width, label='Monthly Target', color='#ff7f0e', alpha=0.8)
                bars2 = ax.bar(x, mtd_target, bar_width, label='MTD Target', color='#2ca02c', alpha=0.8)
                bars3 = ax.bar(x + bar_width, actual, bar_width, label='Actual Disbursed', color='#1f77b4', alpha=0.8)
                
                # Customize the chart
                ax.set_xlabel('Branches')
                ax.set_ylabel('Amount (KSh)')
                ax.set_title('Branch Disbursement vs Targets Comparison')
                ax.set_xticks(x)
                ax.set_xticklabels(branches, rotation=45, ha='right')
                ax.legend()
                
                # Format y-axis to show values in millions
                ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x/1e6:.1f}M'))
                
                # Add value labels on bars
                def add_value_labels(bars):
                    for bar in bars:
                        height = bar.get_height()
                        ax.text(bar.get_x() + bar.get_width()/2., height,
                               f'{height/1e6:.1f}M',
                               ha='center', va='bottom', fontsize=8)
                
                add_value_labels(bars1)
                add_value_labels(bars2)
                add_value_labels(bars3)
                
                # Add MTD achievement percentage bars on the same plot
                # Calculate MTD achievement percentages
                mtd_achievement = (actual / mtd_target * 100).round(1)
                
                # Add percentage text above each actual bar (bars3 is now the actual disbursement)
                for i, (bar, achievement) in enumerate(zip(bars3, mtd_achievement)):
                    height = bar.get_height()
                    ax.text(bar.get_x() + bar.get_width()/2., height + 500000,  # Half a centimeter above (500,000 units)
                           f'{achievement:.1f}%',
                           ha='center', va='bottom', fontsize=9, fontweight='bold', color='#1f77b4')
                
                # Adjust layout and display
                plt.tight_layout()
                st.pyplot(fig)
                
                # Add a comparison table for better readability
                st.subheader("Detailed Comparison Table")
                comparison_table = chart_data.copy()
                comparison_table['Target Gap'] = (comparison_table['Target'] - comparison_table['Total Disbursed']).round(0)
                comparison_table['MTD Gap'] = (comparison_table['MTD Target'] - comparison_table['Total Disbursed']).round(0)
                st.dataframe(comparison_table, use_container_width=True)
                
                # Add target performance metrics
                st.subheader("Target Performance")
                performance_data = chart_data.copy()
                performance_data['Target Achievement %'] = (performance_data['Total Disbursed'] / performance_data['Target'] * 100).round(1)
                performance_data['MTD Achievement %'] = (performance_data['Total Disbursed'] / performance_data['MTD Target'] * 100).round(1)
                
                # Display performance table
                st.dataframe(performance_data[['Total Disbursed', 'Target', 'MTD Target', 'Target Achievement %', 'MTD Achievement %']], use_container_width=True)
            else:
                st.info("No target data available for current branches.")
                st.bar_chart(branch_disbursements['Total Disbursed'])
            
            # Daily Trends
            st.markdown("---")
            st.subheader("📈 Daily Disbursement Trends")
            
            # Convert date column to datetime
            df['Disbursed Date'] = pd.to_datetime(df['Disbursed Date'], format='%d/%m/%Y')
            
            # Daily trend (last 30 days)
            daily_trend = df.groupby(df['Disbursed Date'].dt.date)['Disbursed'].sum().sort_index()
            st.subheader("Daily Disbursement Trend (Last 30 Days)")
            if len(daily_trend) > 0:
                # Get last 30 days or all available days
                recent_days = daily_trend.tail(30)
                st.line_chart(recent_days)
                
                # Daily metrics
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Average Daily", f"{daily_trend.mean():,.0f}")
                with col2:
                    st.metric("Peak Day", f"{daily_trend.max():,.0f}")
                with col3:
                    st.metric("Active Days", len(daily_trend[daily_trend > 0]))
            
        except FileNotFoundError:
            st.error("logbook_disbursements.xlsx file not found in the current directory.")
        except Exception as e:
            st.error(f"Error loading data: {str(e)}")

    with tab2:
        st.subheader("Logbook Collections")
        
        # Load the collections data automatically
        try:
            df_collections = pd.read_csv("logbookrepayments.csv")
            
            # Calculate collections per branch
            branch_collections = df_collections.groupby('branch_id')['repayment_amount'].agg(['sum', 'count', 'mean']).round(2)
            branch_collections.columns = ['Total Collections', 'Number of Repayments', 'Average Repayment']
            
            # Add branch names and replace index
            branch_collections['Branch Name'] = branch_collections.index.map(get_branch_name)
            branch_collections = branch_collections.set_index('Branch Name').sort_values('Total Collections', ascending=False)
            
            # Display summary metrics
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Total Branches", len(branch_collections))
            with col2:
                st.metric("Total Collections", f"{branch_collections['Total Collections'].sum():,.0f}")
            with col3:
                st.metric("Total Repayments", branch_collections['Number of Repayments'].sum())
            with col4:
                st.metric("Average per Repayment", f"{branch_collections['Average Repayment'].mean():,.0f}")
            
            st.markdown("---")
            
            # Display collections per branch
            st.subheader("Collections by Branch")
            st.dataframe(branch_collections, use_container_width=True)
            
            # Create clustered bar chart with targets
            st.subheader("Branch Collection vs Targets Comparison")
            
            # Prepare data for clustered chart
            chart_data = branch_collections[['Total Collections']].copy()
            
            # Add target data for branches that have targets
            chart_data['Collection Target'] = chart_data.index.map(lambda x: LOGBOOK_COLLECTION_TARGETS.get(x, 0))
            
            # Only show branches that have target data
            chart_data = chart_data[chart_data['Collection Target'] > 0]
            
            if not chart_data.empty:
                # Create proper clustered bar chart using matplotlib
                st.subheader("Clustered Bar Chart - Actual vs Targets")
                
                # Prepare data for clustered chart
                branches = chart_data.index.tolist()
                actual = chart_data['Total Collections'].values
                target = chart_data['Collection Target'].values
                
                # Create the clustered bar chart
                fig, ax = plt.subplots(figsize=(12, 6))
                
                # Set the width of bars and positions
                bar_width = 0.35
                x = np.arange(len(branches))
                
                # Create bars
                bars1 = ax.bar(x - bar_width/2, actual, bar_width, label='Actual Collections', color='#1f77b4', alpha=0.8)
                bars2 = ax.bar(x + bar_width/2, target, bar_width, label='Collection Target', color='#ff7f0e', alpha=0.8)
                
                # Remove frame/border
                ax.spines['top'].set_visible(False)
                ax.spines['right'].set_visible(False)
                ax.spines['bottom'].set_visible(False)
                ax.spines['left'].set_visible(False)
                
                # Customize the chart
                ax.set_xlabel('Branches')
                ax.set_ylabel('Amount')
                ax.set_title('Logbook Collections vs Targets Comparison')
                ax.set_xticks(x)
                ax.set_xticklabels(branches, rotation=45, ha='right')
                ax.legend()
                
                # Format y-axis to show values in millions
                ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x/1e6:.1f}M'))
                
                # Add value labels on bars
                def add_value_labels(bars):
                    for bar in bars:
                        height = bar.get_height()
                        ax.text(bar.get_x() + bar.get_width()/2., height,
                               f'{height/1e6:.1f}M',
                               ha='center', va='bottom', fontsize=8)
                
                add_value_labels(bars1)
                add_value_labels(bars2)
                
                # Adjust layout and display
                plt.tight_layout()
                st.pyplot(fig)
                
                # Add a comparison table for better readability
                st.subheader("Detailed Comparison Table")
                comparison_table = chart_data.copy()
                comparison_table['Target Gap'] = (comparison_table['Collection Target'] - comparison_table['Total Collections']).round(0)
                comparison_table['Achievement %'] = (comparison_table['Total Collections'] / comparison_table['Collection Target'] * 100).round(1)
                st.dataframe(comparison_table, use_container_width=True)
            else:
                st.info("No collection target data available for current branches.")
                st.bar_chart(branch_collections['Total Collections'])
            
            # Daily Collections Trends
            st.markdown("---")
            st.subheader("📈 Daily Collection Trends")
            
            # Convert date column to datetime
            df_collections['repayment_collected_date'] = pd.to_datetime(df_collections['repayment_collected_date'], format='%d/%m/%Y')
            
            # Daily trend (last 30 days)
            daily_collections = df_collections.groupby(df_collections['repayment_collected_date'].dt.date)['repayment_amount'].sum().sort_index()
            st.subheader("Daily Collection Trend (Last 30 Days)")
            if len(daily_collections) > 0:
                # Get last 30 days or all available days
                recent_days = daily_collections.tail(30)
                st.line_chart(recent_days)
                
                # Daily metrics
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Average Daily", f"{daily_collections.mean():,.0f}")
                with col2:
                    st.metric("Peak Day", f"{daily_collections.max():,.0f}")
                with col3:
                    st.metric("Active Days", len(daily_collections[daily_collections > 0]))
            
        except FileNotFoundError:
            st.error("logbookrepayments.csv file not found in the current directory.")
        except Exception as e:
            st.error(f"Error loading data: {str(e)}")

    with tab3:
        st.subheader("Portfolio at Risk (PAR)")
        file = st.file_uploader("Upload PAR Excel File", type=["xlsx"])
        if file:
            df = load_excel_data(file)
            st.dataframe(df)

    with tab4:
        st.subheader("Productivity Report")
        file = st.file_uploader("Upload Productivity Excel File", type=["xlsx"])
        if file:
            df = load_excel_data(file)
            st.dataframe(df)

elif menu == "Zidisha":
    tab_dashboard, tab1, tab2, tab3 = st.tabs(["Dashboard", "Disbursements", "Collections", "Productivity"])

    with tab_dashboard:
        st.subheader("Zidisha Dashboard")
        
        # Load data
        try:
            df_zidisha = pd.read_excel("zidisha.xlsx")
        except Exception:
            df_zidisha = pd.DataFrame()

        # Guard: show message if missing
        if df_zidisha.empty:
            st.warning("Missing data: ensure zidisha.xlsx is present.")
        else:
            # Prepare dates
            df_zidisha['Disbursed On Date'] = pd.to_datetime(df_zidisha['Disbursed On Date'], errors='coerce')
            df_zidisha['Expected Matured On Date'] = pd.to_datetime(df_zidisha['Expected Matured On Date'], errors='coerce')
            
            # Use current month and exclude Advans Branch by default
            today = pd.Timestamp.today()
            sel_month = today.strftime('%Y-%m')
            month_start = pd.to_datetime(sel_month + '-01')
            month_end = (month_start + pd.offsets.MonthEnd(0))
            
            # Filter for current month, excluding Advans Branch
            month_mask_disb = (df_zidisha['Disbursed On Date'] >= month_start) & (df_zidisha['Disbursed On Date'] <= month_end)
            month_mask_coll = (df_zidisha['Expected Matured On Date'] >= month_start) & (df_zidisha['Expected Matured On Date'] <= month_end)
            advans_mask = df_zidisha['Branch Name'] != 'Advans Branch'
            
            disb_mtd = df_zidisha[month_mask_disb & advans_mask]
            coll_mtd = df_zidisha[month_mask_coll & advans_mask]

            # KPIs
            total_disb_mtd = float(disb_mtd['Principal Amount'].sum()) if 'Principal Amount' in disb_mtd else 0.0
            total_coll_mtd = float(coll_mtd['Total Repayment Derived'].sum()) if 'Total Repayment Derived' in coll_mtd else 0.0
            total_outstanding = float(df_zidisha[advans_mask]['Total Outstanding Derived'].sum()) if 'Total Outstanding Derived' in df_zidisha else 0.0
            total_expected = float(df_zidisha[advans_mask]['Total Expected Repayment Derived'].sum()) if 'Total Expected Repayment Derived' in df_zidisha else 0.0
            repayment_rate = (total_coll_mtd / total_expected * 100) if total_expected > 0 else 0.0

            # Count unique branches (excluding Advans)
            unique_branches = df_zidisha[df_zidisha['Branch Name'] != 'Advans Branch']['Branch Name'].nunique()
            
            col1, col2, col3, col4, col5 = st.columns(5)
            with col1:
                st.metric("Total Branches", unique_branches)
            with col2:
                st.metric("MTD Disbursed", f"{total_disb_mtd:,.0f}")
            with col3:
                st.metric("MTD Collections", f"{total_coll_mtd:,.0f}")
            with col4:
                st.metric("Total Outstanding", f"{total_outstanding:,.0f}")
            with col5:
                st.metric("Repayment Rate", f"{repayment_rate:.1f}%")

            # Daily Collections vs Disbursements
            st.subheader("Daily Collections vs Disbursements")
            daily_disb = disb_mtd.groupby(disb_mtd['Disbursed On Date'].dt.date)['Principal Amount'].sum().sort_index()
            daily_coll = coll_mtd.groupby(coll_mtd['Expected Matured On Date'].dt.date)['Total Repayment Derived'].sum().sort_index()
            trend_idx = sorted(set(daily_disb.index) | set(daily_coll.index))
            if len(trend_idx) > 0:
                plot_df = pd.DataFrame({
                    'Date': trend_idx,
                    'Disbursed': pd.Series(daily_disb, index=trend_idx).fillna(0),
                    'Collections': pd.Series(daily_coll, index=trend_idx).fillna(0)
                })
                # Use Altair for smooth interpolation
                import altair as alt
                chart = alt.Chart(plot_df).transform_fold(
                    ['Disbursed', 'Collections'],
                    as_=['Metric', 'Value']
                ).mark_line(
                    interpolate='monotone',
                    strokeWidth=2
                ).encode(
                    x=alt.X('Date:T', title='Date'),
                    y=alt.Y('Value:Q', title='Amount'),
                    color=alt.Color('Metric:N', scale=alt.Scale(domain=['Disbursed', 'Collections'], range=['#1f77b4', '#ff7f0e']))
                ).properties(
                    height=300
                )
                st.altair_chart(chart, use_container_width=True)
            else:
                st.info("No disbursements/collections for the selected period.")

            # Branch scatter: Disbursed vs Collection Rate sized by Outstanding
            st.subheader("Branch Performance: Disbursed vs Collection Rate")
            disb_by_branch = disb_mtd.groupby('Branch Name')['Principal Amount'].sum()
            coll_by_branch = coll_mtd.groupby('Branch Name')['Total Repayment Derived'].sum()
            out_by_branch = df_zidisha.groupby('Branch Name')['Total Outstanding Derived'].sum() if 'Total Outstanding Derived' in df_zidisha else pd.Series(dtype=float)
            scatter_df = pd.DataFrame({
                'Disbursed': disb_by_branch,
                'Collections': coll_by_branch,
                'Outstanding': out_by_branch
            }).fillna(0.0)
            # Compute collection rate
            scatter_df['Collection Rate %'] = (scatter_df['Collections'] / scatter_df['Disbursed'].replace({0: np.nan}) * 100).fillna(0)

            if not scatter_df.empty:
                fig, ax = plt.subplots(figsize=(8, 5))
                # Safely scale bubble sizes; handle zero/NaN max outstanding
                max_out = scatter_df['Outstanding'].max()
                try:
                    max_out = float(max_out)
                except Exception:
                    max_out = 0.0
                if max_out and max_out > 0:
                    sizes = (scatter_df['Outstanding'] / max_out * 800).fillna(200)
                else:
                    sizes = pd.Series(200, index=scatter_df.index)
                ax.scatter(scatter_df['Disbursed'], scatter_df['Collection Rate %'], s=sizes, alpha=0.6, c='#1f77b4')
                for name, row in scatter_df.iterrows():
                    ax.text(row['Disbursed'], row['Collection Rate %'], name, fontsize=8, ha='left', va='bottom')
                ax.set_xlabel('MTD Disbursed')
                ax.set_ylabel('Collection Rate %')
                ax.grid(True, alpha=0.3)
                st.pyplot(fig)
            else:
                st.info("No data available for selected filters.")

            # Loan Officer leaderboard
            st.subheader("Top Loan Officers (MTD)")
            if 'Loan Officer Name' in disb_mtd:
                top_officers = disb_mtd.groupby('Loan Officer Name')['Principal Amount'].sum().sort_values(ascending=False).head(10).reset_index()
                top_officers.columns = ['Loan Officer', 'Total Disbursed']
                st.dataframe(top_officers, use_container_width=True)
            else:
                st.info("Loan Officer information not available in disbursements data.")

    with tab1:
        st.subheader("Zidisha Disbursements")
        
        # Load the Zidisha disbursements data automatically
        try:
            from datetime import datetime
            df_zidisha = pd.read_excel("zidisha.xlsx")
            
            # Filter for current month data, excluding Advans Branch
            current_month = datetime.now().month
            current_year = datetime.now().year
            current_month_data = df_zidisha[
                (df_zidisha['Disbursed On Date'].dt.month == current_month) & 
                (df_zidisha['Disbursed On Date'].dt.year == current_year) &
                (df_zidisha['Branch Name'] != 'Advans Branch')
            ]
            
            if not current_month_data.empty:
                # Calculate disbursements per branch for current month
                branch_disbursements = current_month_data.groupby('Branch Name')['Principal Amount'].agg(['sum', 'count', 'mean']).round(2)
                branch_disbursements.columns = ['Total Disbursed', 'Number of Loans', 'Average Disbursement']
                branch_disbursements = branch_disbursements.sort_values('Total Disbursed', ascending=False)
                
                # Display summary metrics
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Total Branches", len(branch_disbursements))
                with col2:
                    st.metric("Total Disbursed", f"{branch_disbursements['Total Disbursed'].sum():,.0f}")
                with col3:
                    st.metric("Total Loans", branch_disbursements['Number of Loans'].sum())
                with col4:
                    st.metric("Average per Loan", f"{branch_disbursements['Average Disbursement'].mean():,.0f}")
                
                st.markdown("---")
                
                # Display disbursements per branch
                st.subheader(f"Disbursements by Branch - {datetime.now().strftime('%B %Y')}")
                st.dataframe(branch_disbursements, use_container_width=True)
                
                # Create a bar chart
                st.subheader("Branch Disbursement Comparison")
                st.bar_chart(branch_disbursements['Total Disbursed'])
                
                # Daily Trends for current month
                st.markdown("---")
                st.subheader("📈 Daily Disbursement Trends (Current Month)")
                
                # Daily trend for current month
                daily_trend = current_month_data.groupby(current_month_data['Disbursed On Date'].dt.date)['Principal Amount'].sum().sort_index()
                st.subheader(f"Daily Disbursement Trend - {datetime.now().strftime('%B %Y')}")
                if len(daily_trend) > 0:
                    st.line_chart(daily_trend)
                    
                    # Daily metrics
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Average Daily", f"{daily_trend.mean():,.0f}")
                    with col2:
                        st.metric("Peak Day", f"{daily_trend.max():,.0f}")
                    with col3:
                        st.metric("Active Days", len(daily_trend[daily_trend > 0]))
            else:
                st.info(f"No disbursement data available for {datetime.now().strftime('%B %Y')}")
                
        except FileNotFoundError:
            st.error("zidisha.xlsx file not found in the current directory.")
        except Exception as e:
            st.error(f"Error loading data: {str(e)}")

    with tab2:
        st.subheader("Zidisha Collections")
        
        # Load the Zidisha collections data automatically
        try:
            from datetime import datetime
            df_zidisha = pd.read_excel("zidisha.xlsx")
            
            # Filter for current month data using Expected Matured On Date, excluding Advans Branch
            current_month = datetime.now().month
            current_year = datetime.now().year
            current_month_data = df_zidisha[
                (df_zidisha['Expected Matured On Date'].dt.month == current_month) & 
                (df_zidisha['Expected Matured On Date'].dt.year == current_year) &
                (df_zidisha['Branch Name'] != 'Advans Branch')
            ]
            
            if not current_month_data.empty:
                # Calculate collections per branch for current month
                branch_collections = current_month_data.groupby('Branch Name')['Total Repayment Derived'].agg(['sum', 'count', 'mean']).round(2)
                branch_collections.columns = ['Total Collections', 'Number of Loans', 'Average Collection']
                branch_collections = branch_collections.sort_values('Total Collections', ascending=False)
                
                # Calculate expected collections for October
                october_expected = df_zidisha[
                    (df_zidisha['Expected Matured On Date'].dt.month == 10) & 
                    (df_zidisha['Expected Matured On Date'].dt.year == 2025) &
                    (df_zidisha['Branch Name'] != 'Advans Branch')
                ]
                total_expected = october_expected['Total Expected Repayment Derived'].sum() if not october_expected.empty else 0
                
                
                # Current Period KPIs (Same Period Last Month)
                st.subheader("📊 Current Period KPIs (Same Period Last Month)")
                
                # Calculate current period and same period last month
                current_date = datetime.now()
                current_month = current_date.month
                current_year = current_date.year
                current_day = current_date.day
                
                # Same period last month
                if current_month == 1:
                    last_month = 12
                    last_year = current_year - 1
                else:
                    last_month = current_month - 1
                    last_year = current_year
                
                # Filter for same period last month (disbursed loans)
                same_period_last_month = df_zidisha[
                    (df_zidisha['Disbursed On Date'].dt.month == last_month) & 
                    (df_zidisha['Disbursed On Date'].dt.year == last_year) &
                    (df_zidisha['Disbursed On Date'].dt.day <= current_day) &
                    (df_zidisha['Branch Name'] != 'Advans Branch')
                ]
                
                if not same_period_last_month.empty:
                    # Calculate KPIs for same period last month
                    total_loans_disbursed = same_period_last_month['Principal Amount'].sum()
                    total_amount_repaid = same_period_last_month['Total Repayment Derived'].sum()
                    total_outstanding = same_period_last_month['Total Outstanding Derived'].sum()
                    total_expected_repayment = same_period_last_month['Total Expected Repayment Derived'].sum()
                    repayment_rate = (total_amount_repaid / total_expected_repayment * 100) if total_expected_repayment > 0 else 0
                    
                    # Display Row 1 KPIs
                    col1, col2, col3, col4, col5 = st.columns(5)
                    with col1:
                        st.metric("Total Loans Disbursed", f"{total_loans_disbursed:,.0f}")
                    with col2:
                        st.metric("Total Amount Repaid", f"{total_amount_repaid:,.0f}")
                    with col3:
                        st.metric("Total Outstanding", f"{total_outstanding:,.0f}")
                    with col4:
                        st.metric("Total Expected Repayment", f"{total_expected_repayment:,.0f}")
                    with col5:
                        st.metric("Repayment Rate", f"{repayment_rate:.1f}%")
                
                st.markdown("---")
                
                # Previous Month Full Month Snapshot
                st.subheader("📈 Previous Month Full Month Snapshot")
                
                # Filter for full previous month (disbursed loans)
                previous_month_full = df_zidisha[
                    (df_zidisha['Disbursed On Date'].dt.month == last_month) & 
                    (df_zidisha['Disbursed On Date'].dt.year == last_year) &
                    (df_zidisha['Branch Name'] != 'Advans Branch')
                ]
                
                if not previous_month_full.empty:
                    # Calculate KPIs for full previous month
                    pm_total_disbursed = previous_month_full['Principal Amount'].sum()
                    pm_total_repaid = previous_month_full['Total Repayment Derived'].sum()
                    pm_total_outstanding = previous_month_full['Total Outstanding Derived'].sum()
                    pm_total_expected = previous_month_full['Total Expected Repayment Derived'].sum()
                    pm_repayment_rate = (pm_total_repaid / pm_total_expected * 100) if pm_total_expected > 0 else 0
                    
                    # Display Row 2 KPIs
                    col1, col2, col3, col4, col5 = st.columns(5)
                    with col1:
                        st.metric(f"{datetime(last_year, last_month, 1).strftime('%b %Y')} — Total Disbursed", f"{pm_total_disbursed:,.0f}")
                    with col2:
                        st.metric(f"{datetime(last_year, last_month, 1).strftime('%b %Y')} — Total Repaid", f"{pm_total_repaid:,.0f}")
                    with col3:
                        st.metric(f"{datetime(last_year, last_month, 1).strftime('%b %Y')} — Total Outstanding", f"{pm_total_outstanding:,.0f}")
                    with col4:
                        st.metric(f"{datetime(last_year, last_month, 1).strftime('%b %Y')} — Total Expected", f"{pm_total_expected:,.0f}")
                    with col5:
                        st.metric(f"{datetime(last_year, last_month, 1).strftime('%b %Y')} — Repayment Rate", f"{pm_repayment_rate:.1f}%")
                
                st.markdown("---")
                
                # Display collections per branch
                st.subheader(f"Collections by Branch - {datetime.now().strftime('%B %Y')}")
                st.dataframe(branch_collections, use_container_width=True)
                
                # Create a bar chart
                st.subheader("Branch Collection Comparison")
                st.bar_chart(branch_collections['Total Collections'])
                
                # Daily Trends for current month
                st.markdown("---")
                st.subheader("📈 Daily Collection Trends (Current Month)")
                
                # Daily trend for current month
                daily_trend = current_month_data.groupby(current_month_data['Expected Matured On Date'].dt.date)['Total Repayment Derived'].sum().sort_index()
                st.subheader(f"Daily Collection Trend - {datetime.now().strftime('%B %Y')}")
                if len(daily_trend) > 0:
                    st.line_chart(daily_trend)
                    
                    # Daily metrics
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Average Daily", f"{daily_trend.mean():,.0f}")
                    with col2:
                        st.metric("Peak Day", f"{daily_trend.max():,.0f}")
                    with col3:
                        st.metric("Active Days", len(daily_trend[daily_trend > 0]))
            else:
                st.info(f"No collection data available for {datetime.now().strftime('%B %Y')}")
                
        except FileNotFoundError:
            st.error("zidisha.xlsx file not found in the current directory.")
        except Exception as e:
            st.error(f"Error loading data: {str(e)}")

    with tab3:
        st.subheader("Zidisha Productivity")
        file = st.file_uploader("Upload Zidisha Productivity Excel File", type=["xlsx"])
        if file:
            df = load_excel_data(file)
            st.dataframe(df)

elif menu == "Kajea - Tech":
    st.info("Upload and visualize Kajea vehicle tracking data here.")
    file = st.file_uploader("Upload Kajea Tracking Excel File", type=["xlsx"])
    if file:
        df = load_excel_data(file)
        st.dataframe(df)

elif menu == "Insurance":
    st.header("🛡️ Insurance Reports")
    st.info("Upload and analyze insurance lead and policy data here.")
    file = st.file_uploader("Upload Insurance Excel File", type=["xlsx"])
    if file:
        df = load_excel_data(file)
        st.dataframe(df)

elif menu == "Advans":
    tab1, tab2 = st.tabs(["Disbursements", "Collections"])

    with tab1:
        st.subheader("Advans Disbursements")
        
        # Load the Advans disbursements data from Zidisha file
        try:
            from datetime import datetime
            df_zidisha = pd.read_excel("zidisha.xlsx")
            
            # Filter for Advans Branch data for current month
            current_month = datetime.now().month
            current_year = datetime.now().year
            advans_data = df_zidisha[
                (df_zidisha['Branch Name'] == 'Advans Branch') &
                (df_zidisha['Disbursed On Date'].dt.month == current_month) & 
                (df_zidisha['Disbursed On Date'].dt.year == current_year)
            ]
            
            if not advans_data.empty:
                # Calculate disbursements for Advans Branch
                total_disbursed = advans_data['Principal Amount'].sum()
                total_loans = len(advans_data)
                average_disbursement = advans_data['Principal Amount'].mean()
                
                # Display summary metrics
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Branch", "Advans Branch")
                with col2:
                    st.metric("Total Disbursed", f"{total_disbursed:,.0f}")
                with col3:
                    st.metric("Total Loans", total_loans)
                with col4:
                    st.metric("Average per Loan", f"{average_disbursement:,.0f}")
                
                st.markdown("---")
                
                # Display detailed loan data
                st.subheader(f"Advans Branch Disbursements - {datetime.now().strftime('%B %Y')}")
                display_data = advans_data[['Client Name', 'Principal Amount', 'Disbursed On Date', 'Loan Officer Name', 'Product Name']].copy()
                display_data = display_data.sort_values('Principal Amount', ascending=False)
                st.dataframe(display_data, use_container_width=True)
                
                # Daily Trends for current month
                st.markdown("---")
                st.subheader("📈 Daily Disbursement Trends (Current Month)")
                
                # Daily trend for current month
                daily_trend = advans_data.groupby(advans_data['Disbursed On Date'].dt.date)['Principal Amount'].sum().sort_index()
                st.subheader(f"Daily Disbursement Trend - {datetime.now().strftime('%B %Y')}")
                if len(daily_trend) > 0:
                    st.line_chart(daily_trend)
                    
                    # Daily metrics
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Average Daily", f"{daily_trend.mean():,.0f}")
                    with col2:
                        st.metric("Peak Day", f"{daily_trend.max():,.0f}")
                    with col3:
                        st.metric("Active Days", len(daily_trend[daily_trend > 0]))
            else:
                st.info(f"No Advans Branch disbursement data available for {datetime.now().strftime('%B %Y')}")
                
        except FileNotFoundError:
            st.error("zidisha.xlsx file not found in the current directory.")
        except Exception as e:
            st.error(f"Error loading data: {str(e)}")

    with tab2:
        st.subheader("Advans Collections")
        
        # Load the Advans collections data from Zidisha file
        try:
            from datetime import datetime
            df_zidisha = pd.read_excel("zidisha.xlsx")
            
            # Filter for Advans Branch data for current month using Expected Matured On Date
            current_month = datetime.now().month
            current_year = datetime.now().year
            advans_data = df_zidisha[
                (df_zidisha['Branch Name'] == 'Advans Branch') &
                (df_zidisha['Expected Matured On Date'].dt.month == current_month) & 
                (df_zidisha['Expected Matured On Date'].dt.year == current_year)
            ]
            
            if not advans_data.empty:
                # Calculate collections for Advans Branch
                total_collections = advans_data['Total Repayment Derived'].sum()
                total_loans = len(advans_data)
                average_collection = advans_data['Total Repayment Derived'].mean()
                
                # Display summary metrics
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Branch", "Advans Branch")
                with col2:
                    st.metric("Total Collections", f"{total_collections:,.0f}")
                with col3:
                    st.metric("Total Loans", total_loans)
                with col4:
                    st.metric("Average per Loan", f"{average_collection:,.0f}")
                
                st.markdown("---")
                
                # Display detailed loan data
                st.subheader(f"Advans Branch Collections - {datetime.now().strftime('%B %Y')}")
                display_data = advans_data[['Client Name', 'Total Repayment Derived', 'Expected Matured On Date', 'Loan Officer Name', 'Product Name']].copy()
                display_data = display_data.sort_values('Total Repayment Derived', ascending=False)
                st.dataframe(display_data, use_container_width=True)
                
                # Daily Trends for current month
                st.markdown("---")
                st.subheader("📈 Daily Collection Trends (Current Month)")
                
                # Daily trend for current month
                daily_trend = advans_data.groupby(advans_data['Expected Matured On Date'].dt.date)['Total Repayment Derived'].sum().sort_index()
                st.subheader(f"Daily Collection Trend - {datetime.now().strftime('%B %Y')}")
                if len(daily_trend) > 0:
                    st.line_chart(daily_trend)
                    
                    # Daily metrics
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Average Daily", f"{daily_trend.mean():,.0f}")
                    with col2:
                        st.metric("Peak Day", f"{daily_trend.max():,.0f}")
                    with col3:
                        st.metric("Active Days", len(daily_trend[daily_trend > 0]))
            else:
                st.info(f"No Advans Branch collection data available for {datetime.now().strftime('%B %Y')}")
                
        except FileNotFoundError:
            st.error("zidisha.xlsx file not found in the current directory.")
        except Exception as e:
            st.error(f"Error loading data: {str(e)}")

# --- FOOTER ---
st.markdown("<hr>", unsafe_allow_html=True)
st.caption("© 2025 Phoenix Capital | Exco Report App")

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

# --- APP TITLE ---
st.title("Exco Report App")

# --- SIDEBAR NAVIGATION ---
st.sidebar.title("📊 Business Units")

# Create distinct vertical navigation buttons
if st.sidebar.button("🚗 Logbook", use_container_width=True, type="primary" if 'menu' not in st.session_state or st.session_state.menu == "Logbook" else "secondary"):
    st.session_state.menu = "Logbook"

if st.sidebar.button("💳 Zidisha", use_container_width=True, type="primary" if 'menu' in st.session_state and st.session_state.menu == "Zidisha" else "secondary"):
    st.session_state.menu = "Zidisha"

if st.sidebar.button("🚚 Kajea - Tracking", use_container_width=True, type="primary" if 'menu' in st.session_state and st.session_state.menu == "Kajea - Tracking" else "secondary"):
    st.session_state.menu = "Kajea - Tracking"

if st.sidebar.button("🛡️ Insurance", use_container_width=True, type="primary" if 'menu' in st.session_state and st.session_state.menu == "Insurance" else "secondary"):
    st.session_state.menu = "Insurance"

if st.sidebar.button("🏦 Advans", use_container_width=True, type="primary" if 'menu' in st.session_state and st.session_state.menu == "Advans" else "secondary"):
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
    st.header("🚗 Logbook Reports")
    tab1, tab2, tab3, tab4 = st.tabs(["Disbursements", "Collections", "PAR", "Productivity"])

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
    st.header("💳 Zidisha Reports")
    tab1, tab2, tab3 = st.tabs(["Disbursements", "Collections", "Productivity"])

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

elif menu == "Kajea - Tracking":
    st.header("🚚 Kajea - Tracking Reports")
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
    st.header("🏦 Advans Reports")
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

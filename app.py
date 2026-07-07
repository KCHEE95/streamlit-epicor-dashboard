import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
import google.generativeai as genai
from datetime import datetime, timedelta

# ---------- Page Config ----------
st.set_page_config(page_title="Epicor Production Dashboard", layout="wide")
st.title("📊 Epicor Kinetic Production Order Dashboard")
st.markdown("Upload Excel report (JobStatByCust3) to generate visual summaries")

# ---------- Gemini AI ----------
def init_gemini(api_key):
    try:
        genai.configure(api_key=api_key)
        return genai.GenerativeModel('gemini-1.5-flash')  # 或 'gemini-1.5-pro'
    except Exception as e:
        st.error(f"Gemini init failed: {e}")
        return None

# ---------- Data Load & Clean ----------
@st.cache_data
def load_data(uploaded_file):
    df = pd.read_excel(uploaded_file, header=0)
    df = df.dropna(how='all')
    if 'Main Part Num' in df.columns:
        df['Main Part Num'] = df['Main Part Num'].ffill()
    else:
        st.error("Column 'Main Part Num' not found")
        return pd.DataFrame()
    if 'Subpart Part Num' not in df.columns:
        st.error("Column 'Subpart Part Num' not found")
        return pd.DataFrame()
    df_sub = df[df['Subpart Part Num'].notna()].copy()
    if df_sub.empty:
        st.warning("No subpart records found")
        return df_sub
    
    # Process Steps
    step_cols = [f'Step {i}' for i in range(1, 21)]
    for col in step_cols:
        if col not in df_sub.columns:
            df_sub[col] = None
    def extract_steps(row):
        steps = []
        for col in step_cols:
            val = row[col]
            if pd.notna(val) and str(val).strip() != '':
                steps.append(str(val).strip())
        return steps
    df_sub['Process Steps'] = df_sub.apply(extract_steps, axis=1)
    df_sub['Step Count'] = df_sub['Process Steps'].apply(len)
    df_sub['First Step'] = df_sub['Process Steps'].apply(lambda x: x[0] if x else None)
    df_sub['Last Step'] = df_sub['Process Steps'].apply(lambda x: x[-1] if x else None)
    
    # Convert dates
    df_sub['First Process Plan Date'] = pd.to_datetime(df_sub['First Process Plan Date'], errors='coerce')
    df_sub['Exwork Date'] = pd.to_datetime(df_sub['Exwork Date'], errors='coerce')
    
    # Define Statuses
    df_sub['Has Plan'] = df_sub['First Process Plan Date'].notna()
    df_sub['Has Current Op'] = df_sub['Current Operation'].notna() & (df_sub['Current Operation'].astype(str).str.strip() != '')
    df_sub['Status'] = 'Pending'
    df_sub.loc[df_sub['Has Plan'] & df_sub['Has Current Op'], 'Status'] = 'Active WIP'
    df_sub.loc[df_sub['Has Plan'] & ~df_sub['Has Current Op'], 'Status'] = 'Completed'
    
    return df_sub

# ---------- Process Summary Function ----------
def get_process_summary(df):
    all_processes = set()
    for steps in df['Process Steps']:
        all_processes.update(steps)
    all_processes = sorted(all_processes)
    
    summary_rows = []
    for proc in all_processes:
        total = 0
        completed = 0
        pending = 0
        for idx, row in df.iterrows():
            steps = row['Process Steps']
            if proc in steps:
                total += 1
                current_op = row['Current Operation']
                if pd.isna(current_op) or str(current_op).strip() == '':
                    completed += 1
                else:
                    try:
                        proc_idx = steps.index(proc)
                    except ValueError:
                        pending += 1
                        continue
                    try:
                        cur_idx = steps.index(current_op)
                    except ValueError:
                        pending += 1
                    else:
                        if proc_idx < cur_idx:
                            completed += 1
                        else:
                            pending += 1
        completion_rate = (completed / total * 100) if total > 0 else 0
        summary_rows.append({
            'Process': proc,
            'Total Subparts': total,
            'Completed': completed,
            'Pending': pending,
            'Completion %': round(completion_rate, 1)
        })
    return pd.DataFrame(summary_rows)

# ---------- Main App ----------
uploaded_file = st.file_uploader("Upload Excel file (.xlsx)", type=['xlsx'])

if uploaded_file is not None:
    df = load_data(uploaded_file)
    if df.empty:
        st.stop()
    st.success(f"✅ Data loaded: {len(df)} subpart records")
    
    # Sidebar for Gemini API Key and Export
    with st.sidebar:
        st.header("🤖 Gemini AI Insights")
        api_key = st.text_input("Enter Google Gemini API Key", type="password")
        if api_key:
            model = init_gemini(api_key)
        else:
            model = None
            st.info("Enter API key to get AI-generated insights")
        
        st.divider()
        st.header("📥 Export Data")
        if st.button("Download Cleaned Excel"):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Cleaned Data')
            st.download_button(
                label="Click to Download",
                data=output.getvalue(),
                file_name=f"epicor_dashboard_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    
    # ---------- KPI Cards ----------
    total = len(df)
    active = len(df[df['Status'] == 'Active WIP'])
    pending = len(df[df['Status'] == 'Pending'])
    completed = len(df[df['Status'] == 'Completed'])
    progress = (completed / total * 100) if total > 0 else 0
    
    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Total Sub Parts", total)
    col2.metric("Active WIP", active)
    col3.metric("Pending", pending)
    col4.metric("Completed", completed)
    col5.metric("Progress %", f"{progress:.1f}%")
    
    # ---------- Tabs ----------
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "📈 Overview", 
        "🔧 Process Analysis", 
        "📊 Process Summary",
        "📋 Data Details", 
        "🧠 AI Insights",
        "⚠️ Urgency (Exwork)"
    ])
    
    with tab1:
        st.subheader("Subpart Count by Main Part")
        main_counts = df['Main Part Num'].value_counts().reset_index()
        main_counts.columns = ['Main Part', 'Count']
        fig1 = px.bar(main_counts, x='Main Part', y='Count', title="Number of Subparts per Main Part")
        st.plotly_chart(fig1, use_container_width=True)
        
        st.subheader("Top 10 Most Frequent Subparts")
        top_sub = df['Subpart Part Num'].value_counts().head(10).reset_index()
        top_sub.columns = ['Subpart', 'Count']
        fig2 = px.bar(top_sub, x='Subpart', y='Count', title="Top 10 Subparts by Occurrence")
        st.plotly_chart(fig2, use_container_width=True)
        
        st.subheader("Completion Rate by Main Part")
        main_status = df.groupby('Main Part Num')['Status'].value_counts().unstack(fill_value=0)
        main_status['Total'] = main_status.sum(axis=1)
        main_status['Completion %'] = (main_status.get('Completed', 0) / main_status['Total'] * 100).round(1)
        main_status_sorted = main_status.sort_values('Completion %', ascending=False)
        fig3 = px.bar(main_status_sorted, x=main_status_sorted.index, y='Completion %', 
                      title="Completion % per Main Part")
        st.plotly_chart(fig3, use_container_width=True)
    
    with tab2:
        st.subheader("Current Operation Distribution")
        op_counts = df[df['Status'] == 'Active WIP']['Current Operation'].value_counts().reset_index()
        op_counts.columns = ['Operation', 'Count']
        fig4 = px.pie(op_counts, names='Operation', values='Count', title="Active WIP by Current Operation")
        st.plotly_chart(fig4, use_container_width=True)
        
        st.subheader("Number of Process Steps per Subpart")
        step_hist = df['Step Count'].value_counts().sort_index().reset_index()
        step_hist.columns = ['Steps', 'Count']
        fig5 = px.bar(step_hist, x='Steps', y='Count', title="Distribution of Process Steps")
        st.plotly_chart(fig5, use_container_width=True)
        
        col1, col2 = st.columns(2)
        with col1:
            first_steps = df['First Step'].value_counts().head(10).reset_index()
            first_steps.columns = ['Step', 'Count']
            fig6 = px.bar(first_steps, x='Step', y='Count', title="Top 10 First Steps")
            st.plotly_chart(fig6, use_container_width=True)
        with col2:
            last_steps = df['Last Step'].value_counts().head(10).reset_index()
            last_steps.columns = ['Step', 'Count']
            fig7 = px.bar(last_steps, x='Step', y='Count', title="Top 10 Last Steps")
            st.plotly_chart(fig7, use_container_width=True)
    
    with tab3:
        st.subheader("Process Summary Table")
        summary_df = get_process_summary(df)
        if not summary_df.empty:
            st.dataframe(summary_df, use_container_width=True)
            col1, col2 = st.columns(2)
            with col1:
                fig8 = px.bar(summary_df, x='Process', y='Completion %', 
                              title="Completion % by Process", 
                              color='Completion %', color_continuous_scale='Blues')
                st.plotly_chart(fig8, use_container_width=True)
            with col2:
                fig9 = px.bar(summary_df, x='Process', y='Pending', 
                              title="Pending Subparts by Process", 
                              color='Pending', color_continuous_scale='Reds')
                st.plotly_chart(fig9, use_container_width=True)
        else:
            st.info("No process data available.")
    
    with tab4:
        st.subheader("Cleaned Data (Key Columns)")
        display_cols = ['Main Part Num', 'Subpart Part Num', 'Subpart Qty', 'JobNum/Asm', 
                        'Current Operation', 'Status', 'Step Count', 'Process Steps', 
                        'First Step', 'Last Step', 'First Process Plan Date']
        available = [c for c in display_cols if c in df.columns]
        st.dataframe(df[available], use_container_width=True)
    
    with tab5:
        st.subheader("🧠 AI-Generated Business Insights")
        if model and api_key:
            # Prepare summary stats
            summary = {
                "Total Sub Parts": total,
                "Active WIP": active,
                "Pending": pending,
                "Completed": completed,
                "Progress %": progress,
                "Most Common Current Operation": df[df['Status']=='Active WIP']['Current Operation'].mode().iloc[0] if not df[df['Status']=='Active WIP'].empty else None,
                "Average Steps": df['Step Count'].mean(),
                "Main Part with highest completion": main_status['Completion %'].idxmax() if not main_status.empty else None,
            }
            prompt = f"""
            You are a production planning analyst. Based on the following summary statistics from a manufacturing order report, provide concise business insights. Highlight potential bottlenecks, efficiency suggestions, or areas that need attention.
            Summary:
            {summary}
            Please respond in English, using bullet points.
            """
            try:
                response = model.generate_content(prompt)
                st.write(response.text)
            except Exception as e:
                st.error(f"AI call failed: {e}")
        else:
            st.info("Enter a valid Gemini API key in the sidebar to get AI insights.")
    
    with tab6:
        st.subheader("Urgency by Exwork Date")
        if 'Exwork Date' in df.columns:
            df['Exwork Date'] = pd.to_datetime(df['Exwork Date'], errors='coerce')
            today = datetime.now().date()
            def urgency_group(date):
                if pd.isna(date):
                    return "No Date"
                delta = (date.date() - today).days
                if delta < 0:
                    return "Overdue"
                elif delta <= 7:
                    return "Next 7 Days"
                elif delta <= 30:
                    return "Next 30 Days"
                else:
                    return "Beyond 30 Days"
            df['Urgency'] = df['Exwork Date'].apply(urgency_group)
            urgency_counts = df['Urgency'].value_counts().reset_index()
            urgency_counts.columns = ['Urgency', 'Count']
            fig10 = px.bar(urgency_counts, x='Urgency', y='Count', 
                          title="Subparts by Exwork Date Urgency", 
                          color='Urgency', color_discrete_sequence=px.colors.qualitative.Set3)
            st.plotly_chart(fig10, use_container_width=True)
            st.subheader("Overdue or Next 7 Days")
            urgent_df = df[df['Urgency'].isin(['Overdue', 'Next 7 Days'])]
            if not urgent_df.empty:
                st.dataframe(urgent_df[['Subpart Part Num', 'Main Part Num', 'Exwork Date', 'Current Operation', 'Status']])
            else:
                st.info("No urgent items found.")
        else:
            st.info("Column 'Exwork Date' not available in the data.")

else:
    st.info("👈 Please upload an Excel file to start the analysis.")

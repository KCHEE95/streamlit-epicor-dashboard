import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
import google.generativeai as genai
from datetime import datetime

# ---------- 页面配置 ----------
st.set_page_config(page_title="Epicor 订单仪表盘", layout="wide")
st.title("📊 Epicor Kinetic 生产订单分析仪表盘")
st.markdown("上传 Excel 报表 (JobStatByCust3) 自动生成可视化摘要")

# ---------- Gemini AI 配置 ----------
def init_gemini(api_key):
    try:
        genai.configure(api_key=api_key)
        return genai.GenerativeModel('gemini-1.5-flash')
    except Exception as e:
        st.error(f"Gemini 初始化失败: {e}")
        return None

# ---------- 数据加载与清洗 ----------
@st.cache_data
def load_data(uploaded_file):
    # 读取 Excel，保留所有列
    df = pd.read_excel(uploaded_file, header=0)
    
    # 1. 删除全空行（包括所有列为空）
    df = df.dropna(how='all')
    
    # 2. 前向填充 Main Part Num（因每张主零件只有首行有值）
    if 'Main Part Num' in df.columns:
        df['Main Part Num'] = df['Main Part Num'].ffill()   # 修复弃用警告
    else:
        st.error("Excel 中缺少 'Main Part Num' 列，请检查文件格式")
        return pd.DataFrame()
    
    # 3. 只保留有子零件号的行（即实际子零件记录）
    if 'Subpart Part Num' not in df.columns:
        st.error("Excel 中缺少 'Subpart Part Num' 列")
        return pd.DataFrame()
    df_sub = df[df['Subpart Part Num'].notna()].copy()
    
    if df_sub.empty:
        st.warning("未找到任何子零件数据，请确认 Excel 格式")
        return df_sub
    
    # 4. 提取工序步骤（列 Z ~ AS，即 Step 1 ~ Step 20）
    step_cols = [f'Step {i}' for i in range(1, 21)]
    # 确保所有 Step 列都存在，缺失则补空列
    for col in step_cols:
        if col not in df_sub.columns:
            df_sub[col] = None
    
    # 5. 构建每个子零件的工序列表（过滤空值）
    def extract_steps(row):
        steps = []
        for col in step_cols:
            val = row[col]
            # 判断非空：不为 NaN，且去除空格后非空字符串
            if pd.notna(val) and str(val).strip() != '':
                steps.append(str(val).strip())
        return steps
    
    df_sub['Process Steps'] = df_sub.apply(extract_steps, axis=1)
    df_sub['Step Count'] = df_sub['Process Steps'].apply(len)
    df_sub['First Step'] = df_sub['Process Steps'].apply(lambda x: x[0] if x else None)
    df_sub['Last Step'] = df_sub['Process Steps'].apply(lambda x: x[-1] if x else None)
    
    return df_sub

# ---------- 主界面 ----------
uploaded_file = st.file_uploader("上传 Excel 文件 (.xlsx)", type=['xlsx'])

if uploaded_file is not None:
    df = load_data(uploaded_file)
    
    if df.empty:
        st.stop()
    
    st.success(f"✅ 数据加载成功，共 {len(df)} 个子零件行")
    
    # 侧边栏：Gemini API Key 输入
    with st.sidebar:
        st.header("🤖 Gemini AI 智能分析")
        api_key = st.text_input("请输入 Google Gemini API Key", type="password")
        if api_key:
            model = init_gemini(api_key)
        else:
            model = None
            st.info("输入 API Key 可获取 AI 生成的业务洞察")
        
        st.divider()
        st.header("📥 数据导出")
        if st.button("下载处理后的 Excel"):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Cleaned Data')
            st.download_button(
                label="点击下载",
                data=output.getvalue(),
                file_name=f"epicor_dashboard_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    
    # ---------- 仪表盘 Tabs ----------
    tab1, tab2, tab3, tab4 = st.tabs(["📈 总览", "🔧 工序分析", "📋 数据明细", "🧠 AI 洞察"])
    
    with tab1:
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("子零件总数", len(df))
        col2.metric("主零件种类", df['Main Part Num'].nunique())
        col3.metric("不同当前工序", df['Current Operation'].nunique())
        col4.metric("平均工序步数", f"{df['Step Count'].mean():.1f}")
        
        st.subheader("📊 子零件数量按主零件分布")
        main_part_counts = df['Main Part Num'].value_counts().reset_index()
        main_part_counts.columns = ['Main Part', 'Count']
        fig = px.bar(main_part_counts, x='Main Part', y='Count', title="各主零件包含的子零件数")
        st.plotly_chart(fig, use_container_width=True)
        
        st.subheader("📊 子零件数量 Top 10")
        top_parts = df['Subpart Part Num'].value_counts().head(10).reset_index()
        top_parts.columns = ['Subpart', 'Count']
        fig2 = px.bar(top_parts, x='Subpart', y='Count', title="出现频次最高的子零件")
        st.plotly_chart(fig2, use_container_width=True)
    
    with tab2:
        st.subheader("当前工序分布")
        op_counts = df['Current Operation'].value_counts().reset_index()
        op_counts.columns = ['Operation', 'Count']
        fig3 = px.pie(op_counts, names='Operation', values='Count', title="当前工序占比")
        st.plotly_chart(fig3, use_container_width=True)
        
        st.subheader("工序步数分布")
        step_hist = df['Step Count'].value_counts().sort_index().reset_index()
        step_hist.columns = ['Steps', 'Count']
        fig4 = px.bar(step_hist, x='Steps', y='Count', title="每个子零件的工序步数分布")
        st.plotly_chart(fig4, use_container_width=True)
        
        # 首工序 / 末工序 统计
        col1, col2 = st.columns(2)
        with col1:
            first_step_counts = df['First Step'].value_counts().head(10).reset_index()
            first_step_counts.columns = ['Step', 'Count']
            fig5 = px.bar(first_step_counts, x='Step', y='Count', title="最常见的首工序")
            st.plotly_chart(fig5, use_container_width=True)
        with col2:
            last_step_counts = df['Last Step'].value_counts().head(10).reset_index()
            last_step_counts.columns = ['Step', 'Count']
            fig6 = px.bar(last_step_counts, x='Step', y='Count', title="最常见的末工序")
            st.plotly_chart(fig6, use_container_width=True)
    
    with tab3:
        st.subheader("原始数据（清洗后）")
        # 显示关键列
        display_cols = ['Main Part Num', 'Subpart Part Num', 'Subpart Qty', 'JobNum/Asm', 
                        'Current Operation', 'Step Count', 'Process Steps', 'First Step', 'Last Step']
        # 确保这些列存在
        available_cols = [c for c in display_cols if c in df.columns]
        st.dataframe(df[available_cols], use_container_width=True)
    
    with tab4:
        st.subheader("🧠 AI 生成的业务洞察")
        if model is not None and api_key:
            # 准备数据摘要
            summary_stats = {
                "总子零件数": len(df),
                "主零件种类": df['Main Part Num'].nunique(),
                "当前工序分布": df['Current Operation'].value_counts().to_dict(),
                "最常见首工序": df['First Step'].mode().iloc[0] if not df['First Step'].mode().empty else None,
                "最常见末工序": df['Last Step'].mode().iloc[0] if not df['Last Step'].mode().empty else None,
                "平均步数": df['Step Count'].mean(),
                "子零件数量最多的主零件": df['Main Part Num'].value_counts().idxmax(),
            }
            prompt = f"""
            你是一位生产运营分析师。请根据以下数据摘要，提供简明扼要的业务洞察，指出可能的瓶颈、效率提升点或需要关注的地方。
            摘要数据：
            {summary_stats}
            请用中文回答，格式清晰，分点说明。
            """
            try:
                response = model.generate_content(prompt)
                st.write(response.text)
            except Exception as e:
                st.error(f"AI 调用失败: {e}")
        else:
            st.info("请在侧边栏输入有效的 Gemini API Key 来获取 AI 洞察。")
else:
    st.info("👈 请上传 Excel 文件以开始分析")

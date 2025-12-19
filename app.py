import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import numpy as np
from io import BytesIO
import warnings
import base64
warnings.filterwarnings('ignore')

# --------------------------
# 1. 基础配置（保留原功能，新增Plotly中文支持）
# --------------------------
# 兼容残留Matplotlib配置（虽未用Matplotlib绘图，避免潜在冲突）
try:
    import matplotlib.pyplot as plt
    plt.rcParams['font.sans-serif'] = ['SimHei', 'Arial Unicode MS']
    plt.rcParams['axes.unicode_minus'] = False
except ImportError:
    pass  # 若未安装Matplotlib不影响核心功能

st.set_page_config(
    page_title="数字化转型指数分析App",
    page_icon="📊",
    layout="wide"
)

# Plotly中文配置（确保悬停文本、图例等中文正常显示）
plotly_config = {
    'displayModeBar': True,
    'locale': 'zh-CN'
}

# PDF显示函数（保留原功能）
def display_pdf(pdf_data, height=800):
    try:
        if isinstance(pdf_data, str) and pdf_data.endswith(".pdf"):
            with open(pdf_data, "rb") as f:
                pdf_bytes = f.read()
        elif hasattr(pdf_data, "getbuffer"):
            pdf_bytes = pdf_data.getbuffer()
        elif isinstance(pdf_data, (bytes, BytesIO)):
            pdf_bytes = pdf_data if isinstance(pdf_data, bytes) else pdf_data.read()
        else:
            st.error("❌ 不支持的PDF数据类型，请传入本地路径、上传文件或字节流")
            return
        
        base64_pdf = base64.b64encode(pdf_bytes).decode('utf-8')
        pdf_display = f"""
        <iframe 
            src="data:application/pdf;base64,{base64_pdf}" 
            width="100%" 
            height="{height}" 
            type="application/pdf"
            style="border: none; border-radius: 4px;"
        ></iframe>
        """
        st.markdown(pdf_display, unsafe_allow_html=True)
    except Exception as e:
        st.error(f"❌ PDF显示失败：{str(e)}")

# --------------------------
# 2. 数据读取与清洗（固定Excel路径为C:\Users\张珊\Desktop\3\数字化转型指数汇总_行业信息完整.xlsx）
# --------------------------
@st.cache_data  # 缓存数据，避免重复读取
def load_data():
    # 固定Excel文件路径（已按要求设置为目标路径）
    excel_path = r"C:\Users\张珊\Desktop\3\数字化转型指数汇总_行业信息完整.xlsx"
    try:
        # 读取Excel文件（指定openpyxl引擎，确保.xlsx文件兼容）
        df = pd.read_excel(excel_path, sheet_name="Sheet1", engine="openpyxl")
    except FileNotFoundError:
        st.error(f"❌ 未找到Excel文件，请检查路径：{excel_path}")
        st.stop()  # 停止运行，避免后续报错
    except Exception as e:
        st.error(f"❌ Excel文件读取失败：{str(e)}（可能是文件损坏或格式不兼容，建议用Excel打开确认）")
        st.stop()
    
    # 1. 校验必要字段（与Excel列名完全匹配）
    required_columns = ["股票代码", "企业名称", "年份", "数字化转型综合指数", "行业代码", "行业名称"]
    missing_cols = [col for col in required_columns if col not in df.columns]
    if missing_cols:
        st.error(f"❌ Excel表缺少必要字段：{', '.join(missing_cols)}")
        st.stop()
    
    # 2. 数据清洗（删除空值、重复值，规范字段类型）
    df_clean = df[required_columns].copy()
    # 删除关键字段为空的行，重置索引避免筛选错位
    df_clean = df_clean.dropna(subset=required_columns).reset_index(drop=True)
    # 规范数据类型：年份→整数（排除异常值），指数→数值型
    df_clean["年份"] = pd.to_numeric(df_clean["年份"], errors="coerce")
    df_clean = df_clean[df_clean["年份"].notna()].reset_index(drop=True)  # 移除年份为空的异常行
    df_clean["年份"] = df_clean["年份"].astype(int)
    df_clean["数字化转型综合指数"] = pd.to_numeric(df_clean["数字化转型综合指数"], errors="coerce")
    # 删除重复行，再次重置索引
    df_clean = df_clean.drop_duplicates().reset_index(drop=True)
    
    # 3. 重命名字段（与后续功能逻辑统一）
    df_clean.rename(columns={
        "股票代码": "企业代码",
        "数字化转型综合指数": "数字化转型指数"
    }, inplace=True)
    
    # 4. 计算行业平均指数（按行业+年份分组，避免重复计算）
    industry_avg = df_clean.groupby(["行业代码", "行业名称", "年份"])["数字化转型指数"].mean().reset_index()
    industry_avg.rename(columns={"数字化转型指数": "行业平均指数"}, inplace=True)
    
    return df_clean, industry_avg

# 读取数据（调用固定路径的加载函数）
enterprise_data, industry_avg = load_data()

# --------------------------
# 3. Plotly交互图表生成函数（核心悬停功能，保留原逻辑）
# --------------------------
def create_hover_chart(x_data, y_data_list, labels, title, x_label="年份", y_label="数字化转型指数"):
    """
    生成支持鼠标悬停的Plotly折线图
    - x_data: X轴数据（年份，统一数组确保对齐）
    - y_data_list: Y轴数据列表（如[企业指数数组, 行业平均指数数组]）
    - labels: 每条折线的名称（如["平安银行指数", "货币金融服务平均指数"]）
    - title: 图表标题
    """
    fig = go.Figure()
    # 遍历所有Y轴数据，添加折线（显示线+点，确保悬停可触发）
    for y_data, label in zip(y_data_list, labels):
        fig.add_trace(go.Scatter(
            x=x_data,
            y=y_data,
            mode="lines+markers",
            name=label,
            # 悬停文本：自定义显示“年份+数值”（保留4位小数，提升精度）
            hovertemplate=f"{x_label}：%{{x}}<br>{label}：%{{y:.4f}}<extra></extra>",
            line=dict(width=2.5),
            marker=dict(size=6)  # 点放大，便于鼠标捕捉
        ))
    
    # 图表样式配置（优化中文显示与布局）
    fig.update_layout(
        title=title,
        xaxis_title=x_label,
        yaxis_title=y_label,
        hovermode="closest",  # 鼠标靠近点时优先显示该点数据，避免多线干扰
        width=1200,
        height=600,
        legend=dict(x=0.01, y=0.99, bgcolor="rgba(255,255,255,0.8)"),  # 图例放在左上角，半透明背景
        font=dict(family="SimHei", size=12)  # 全局字体设置为黑体，避免中文乱码
    )
    return fig

# --------------------------
# 4. 侧边栏导航（保留原功能，优化数据概览显示）
# --------------------------
st.sidebar.title("📋 功能导航")
# 核心查询类型选择
query_type = st.sidebar.radio(
    "请选择查询类型",
    ["企业数字化指数查询", "行业数字化指数查询", "多行业对比分析", "PDF报告预览"],
    index=0  # 默认选中第一个功能
)

# PDF上传配置（保留原功能，适配本地路径与上传文件）
pdf_file = None
if query_type == "PDF报告预览":
    st.sidebar.divider()
    st.sidebar.subheader("📄 PDF文件来源")
    pdf_source = st.sidebar.radio("选择PDF来源", ["本地文件路径", "上传PDF文件"], index=1)
    
    if pdf_source == "本地文件路径":
        pdf_local_path = st.sidebar.text_input(
            "输入PDF本地路径", 
            placeholder=r"示例：C:\Users\XXX\Desktop\报告.pdf",
            help="若路径包含中文，直接输入即可"
        )
        if pdf_local_path:
            pdf_file = pdf_local_path  # 赋值为本地路径
    else:
        pdf_uploaded = st.sidebar.file_uploader("选择PDF文件", type="pdf", help="支持最大100MB的PDF文件")
        if pdf_uploaded:
            pdf_file = pdf_uploaded  # 赋值为上传文件对象

# 数据概览（优化显示逻辑，避免数据异常）
st.sidebar.divider()
st.sidebar.subheader("📊 数据概览")
try:
    enterprise_count = enterprise_data["企业名称"].nunique()
    industry_count = industry_avg["行业名称"].nunique()
    year_min = enterprise_data["年份"].min()
    year_max = enterprise_data["年份"].max()
    st.sidebar.write(f"企业数量：{enterprise_count} 家")
    st.sidebar.write(f"行业数量：{industry_count} 个")
    st.sidebar.write(f"数据年份范围：{year_min} - {year_max}")
except Exception as e:
    st.sidebar.warning(f"⚠️ 数据概览加载失败：{str(e)}")

# --------------------------
# 5. 核心功能1：企业数字化指数查询（集成Plotly交互）
# --------------------------
if query_type == "企业数字化指数查询":
    st.title("🏢 企业数字化转型指数查询")
    st.divider()
    
    # 双输入框：支持企业代码/名称模糊查询（带示例提示）
    col1, col2 = st.columns(2)
    with col1:
        enterprise_code = st.text_input("输入企业代码（如：000820）", placeholder="支持模糊匹配，例：0008")
    with col2:
        enterprise_name = st.text_input("输入企业名称（如：平安银行）", placeholder="支持模糊匹配，例：平安")
    
    # 触发查询逻辑（任一输入框有内容即执行查询）
    if enterprise_code or enterprise_name:
        # 初始化筛选掩码（避免索引不匹配导致的筛选错误）
        filter_mask = np.zeros(len(enterprise_data), dtype=bool)
        # 企业代码筛选（转为字符串避免数值匹配误差，如000820被识别为820）
        if enterprise_code:
            filter_mask |= enterprise_data["企业代码"].astype(str).str.contains(enterprise_code, case=False, na=False)
        # 企业名称筛选（不区分大小写，忽略空值）
        if enterprise_name:
            filter_mask |= enterprise_data["企业名称"].str.contains(enterprise_name, case=False, na=False)
        
        # 筛选结果排序，重置索引
        result = enterprise_data[filter_mask].sort_values(["企业名称", "年份"]).reset_index(drop=True)
        
        # 处理无匹配结果的情况
        if result.empty:
            st.warning("⚠️ 未找到匹配的企业，请检查输入关键词（如特殊字符*ST需完整输入）或尝试其他查询方式")
        else:
            # 多企业匹配时，让用户选择具体企业（避免数据混淆）
            unique_enterprises = result[["企业代码", "企业名称"]].drop_duplicates().reset_index(drop=True)
            if len(unique_enterprises) > 1:
                st.subheader("🔍 匹配到以下企业，请选择目标企业")
                selected_enterprise = st.selectbox(
                    "选择企业",
                    options=unique_enterprises.apply(lambda x: f"{x['企业名称']}（代码：{x['企业代码']}）", axis=1),
                    help="若企业名称重复，可通过代码区分"
                )
                # 提取选中企业的名称与代码
                selected_name = selected_enterprise.split("（代码：")[0]
                selected_code = selected_enterprise.split("（代码：")[1].replace("）", "")
                # 筛选该企业的详细数据（按年份排序）
                enterprise_detail = result[
                    (result["企业名称"] == selected_name) & 
                    (result["企业代码"] == selected_code)
                ].sort_values("年份").reset_index(drop=True)
            else:
                # 仅匹配到1家企业，直接提取数据
                selected_name = unique_enterprises.iloc[0]["企业名称"]
                selected_code = unique_enterprises.iloc[0]["企业代码"]
                enterprise_detail = result.sort_values("年份").reset_index(drop=True)
            
            # 1. 显示企业基础信息（行业、数据时间范围）
            st.subheader(f"📈 {selected_name}（代码：{selected_code}）数字化转型指数")
            industry_info = enterprise_detail.iloc[0][["行业代码", "行业名称"]]
            st.write(f"所属行业：{industry_info['行业名称']}（行业代码：{industry_info['行业代码']}）")
            st.write(f"数据时间范围：{enterprise_detail['年份'].min()} - {enterprise_detail['年份'].max()}")
            
            # 2. 匹配行业平均数据（确保年份对齐，避免部分年份缺失导致图表错位）
            industry_index = industry_avg[
                (industry_avg["行业代码"] == industry_info["行业代码"]) & 
                (industry_avg["年份"].isin(enterprise_detail["年份"]))
            ].sort_values("年份").reset_index(drop=True)
            # 强制对齐年份（用企业数据的年份为基准，补全行业平均数据）
            merged_years = enterprise_detail["年份"].values
            industry_index_aligned = industry_index.set_index("年份").reindex(merged_years).reset_index()["行业平均指数"].values
            
            # 3. 生成并显示交互图表（核心功能：鼠标悬停显示数值）
            st.subheader("📊 指数趋势图（鼠标悬停查看具体数值）")
            fig = create_hover_chart(
                x_data=merged_years,
                y_data_list=[
                    enterprise_detail["数字化转型指数"].values,
                    industry_index_aligned
                ],
                labels=[f"{selected_name}指数", f"{industry_info['行业名称']}平均指数"],
                title=f"{selected_name}数字化转型指数趋势（{merged_years.min()}-{merged_years.max()}）"
            )
            # 显示图表（适配页面宽度，传递中文配置）
            st.plotly_chart(fig, use_container_width=True, config=plotly_config)
            
            # 4. 显示历年详细数据表格（重命名列名，提升可读性）
            st.subheader("📋 历年详细数据")
            display_cols = ["年份", "数字化转型指数", "行业代码", "行业名称"]
            st.dataframe(
                enterprise_detail[display_cols].rename(columns={"数字化转型指数": "数字化转型综合指数"}),
                use_container_width=True,
                hide_index=True  # 隐藏索引列，避免用户混淆
            )

# --------------------------
# 6. 核心功能2：行业数字化指数查询（集成Plotly交互）
# --------------------------
elif query_type == "行业数字化指数查询":
    st.title("🏭 行业数字化转型指数查询")
    st.divider()
    
    # 双输入框：支持行业代码/名称模糊查询
    col1, col2 = st.columns(2)
    with col1:
        industry_code = st.text_input("输入行业代码（如：J66）", placeholder="支持模糊匹配，例：J")
    with col2:
        industry_name = st.text_input("输入行业名称（如：货币金融服务）", placeholder="支持模糊匹配，例：金融")
    
    # 触发查询逻辑
    if industry_code or industry_name:
        # 初始化筛选掩码
        filter_mask = np.zeros(len(industry_avg), dtype=bool)
        # 行业代码筛选（不区分大小写）
        if industry_code:
            filter_mask |= industry_avg["行业代码"].str.contains(industry_code, case=False, na=False)
        # 行业名称筛选（不区分大小写）
        if industry_name:
            filter_mask |= industry_avg["行业名称"].str.contains(industry_name, case=False, na=False)
        
        # 筛选结果排序，重置索引
        result = industry_avg[filter_mask].sort_values(["行业名称", "年份"]).reset_index(drop=True)
        
        # 处理无匹配结果的情况
        if result.empty:
            st.warning("⚠️ 未找到匹配的行业，请检查输入关键词（如行业名称是否包含特殊符号）")
        else:
            # 多行业匹配时，让用户选择具体行业
            unique_industries = result[["行业代码", "行业名称"]].drop_duplicates().reset_index(drop=True)
            if len(unique_industries) > 1:
                st.subheader("🔍 匹配到以下行业，请选择目标行业")
                selected_industry = st.selectbox(
                    "选择行业",
                    options=unique_industries.apply(lambda x: f"{x['行业名称']}（代码：{x['行业代码']}）", axis=1)
                )
                # 提取选中行业的名称与代码
                selected_ind_name = selected_industry.split("（代码：")[0]
                selected_ind_code = selected_industry.split("（代码：")[1].replace("）", "")
                # 筛选该行业的详细数据
                industry_detail = result[
                    (result["行业名称"] == selected_ind_name) & 
                    (result["行业代码"] == selected_ind_code)
                ].sort_values("年份").reset_index(drop=True)
            else:
                # 仅匹配到1个行业，直接提取数据
                selected_ind_name = unique_industries.iloc[0]["行业名称"]
                selected_ind_code = unique_industries.iloc[0]["行业代码"]
                industry_detail = result.sort_values("年份").reset_index(drop=True)
            
            # 1. 显示行业基础信息
            st.subheader(f"📈 {selected_ind_name}（代码：{selected_ind_code}）数字化转型指数")
            st.write(f"数据时间范围：{industry_detail['年份'].min()} - {industry_detail['年份'].max()}")
            
            # 2. 生成并显示交互图表
            st.subheader("📊 行业平均指数趋势图（鼠标悬停查看具体数值）")
            fig = create_hover_chart(
                x_data=industry_detail["年份"].values,
                y_data_list=[industry_detail["行业平均指数"].values],
                labels=[f"{selected_ind_name}平均指数"],
                title=f"{selected_ind_name}数字化转型平均指数趋势（{industry_detail['年份'].min()}-{industry_detail['年份'].max()}）"
            )
            st.plotly_chart(fig, use_container_width=True, config=plotly_config)
            
            # 3. 显示行业历年平均数据表格
            st.subheader("📋 历年行业平均指数")
            st.dataframe(
                industry_detail[["年份", "行业平均指数"]].rename(columns={"行业平均指数": "数字化转型平均指数"}),
                use_container_width=True,
                hide_index=True
            )

# --------------------------
# 7. 核心功能3：多行业对比分析（集成Plotly交互）
# --------------------------
elif query_type == "多行业对比分析":
    st.title("📊 多行业数字化转型指数对比")
    st.divider()
    st.write("💡 选择多个行业，对比其数字化转型指数趋势（含全选行业平均线）")
    
    # 行业选择：下拉多选，带搜索功能（按行业名称排序，优化选择体验）
    all_industries = industry_avg[["行业代码", "行业名称"]].drop_duplicates().sort_values("行业名称").reset_index(drop=True)
    selected_industries = st.multiselect(
        "请选择要对比的行业（可多选，建议3-5个）",
        options=all_industries.apply(lambda x: f"{x['行业名称']}（代码：{x['行业代码']}）", axis=1),
        default=all_industries.apply(lambda x: f"{x['行业名称']}（代码：{x['行业代码']}）", axis=1).head(2),  # 默认选前2个行业
        help="选择过多行业会导致图表拥挤，建议不超过5个"
    )
    
    # 当选择行业数量≥1时，执行对比逻辑
    if selected_industries:
        # 提取选中行业的名称与代码
        selected_ind_names = [ind.split("（代码：")[0] for ind in selected_industries]
        selected_ind_codes = [ind.split("（代码：")[1].replace("）", "") for ind in selected_industries]
        
        # 筛选选中行业的平均指数数据
        compare_data = industry_avg[industry_avg["行业名称"].isin(selected_ind_names)].sort_values(["行业名称", "年份"]).reset_index(drop=True)
        # 计算全选行业的整体平均指数（用于对比参考）
        overall_avg = compare_data.groupby("年份")["行业平均指数"].mean().reset_index()
        overall_avg.rename(columns={"行业平均指数": "全选行业平均指数"}, inplace=True)
        
        # 1. 显示对比行业的基础信息（名称+代码）
        st.subheader("🔍 对比行业信息")
        st.dataframe(
            pd.DataFrame({
                "行业名称": selected_ind_names,
                "行业代码": selected_ind_codes
            }),
            use_container_width=True,
            hide_index=True
        )
        
        # 2. 准备对比数据（确保所有行业年份对齐，避免图表错位）
        all_years = compare_data["年份"].unique()
        y_data_list = []
        labels = []
        # 遍历每个行业，按统一年份对齐数据
        for industry in selected_ind_names:
            ind_data = compare_data[compare_data["行业名称"] == industry].set_index("年份").reindex(all_years).reset_index()
            y_data_list.append(ind_data["行业平均指数"].values)
            labels.append(f"{industry}平均指数")
        # 添加全选行业平均线（按统一年份对齐）
        overall_avg_aligned = overall_avg.set_index("年份").reindex(all_years).reset_index()["全选行业平均指数"].values
        y_data_list.append(overall_avg_aligned)
        labels.append("全选行业平均指数")
        
        # 3. 生成多行业交互对比图
        st.subheader("📈 多行业指数对比趋势图（鼠标悬停查看具体数值）")
        fig = create_hover_chart(
            x_data=all_years,
            y_data_list=y_data_list,
            labels=labels,
            title="多行业数字化转型指数对比分析"
        )
        st.plotly_chart(fig, use_container_width=True, config=plotly_config)
        
        # 4. 显示多行业历年数据对比表（透视表格式，更直观）
        st.subheader("📋 多行业历年指数对比表")
        compare_table = compare_data.pivot_table(
            index="年份",
            columns="行业名称",
            values="行业平均指数",
            fill_value="-"  # 空值用"-"填充，避免显示NaN
        ).round(4)  # 保留4位小数，提升精度
        # 添加全选行业平均列（最后一列，便于对比）
        compare_table["全选行业平均指数"] = overall_avg.set_index("年份")["全选行业平均指数"].round(4)
        st.dataframe(compare_table, use_container_width=True)

# --------------------------
# 8. 功能4：PDF报告预览（保留原功能，优化错误处理）
# --------------------------
elif query_type == "PDF报告预览":
    st.title("📄 数字化转型指数PDF报告预览")
    st.divider()
    st.write("💡 支持上传本地PDF报告或输入文件路径，在线预览报告内容（无需下载）")
    
    # 显示PDF文件（根据侧边栏选择的来源）
    if pdf_file:
        st.subheader("📖 报告预览")
        # 调用PDF显示函数，设置高度为800px（适配大多数报告）
        display_pdf(pdf_file, height=800)
        
        # 显示PDF文件信息（大小、名称）
        try:
            if hasattr(pdf_file, "name"):  # 上传文件对象
                file_name = pdf_file.name
                file_size = f"{pdf_file.size / (1024*1024):.2f} MB"  # 转换为MB
            else:  # 本地文件路径
                import os
                file_name = os.path.basename(pdf_file)
                file_size = f"{os.path.getsize(pdf_file) / (1024*1024):.2f} MB"
            
            st.subheader("📊 文件信息")
            st.dataframe(
                pd.DataFrame({
                    "文件属性": ["文件名称", "文件大小", "预览方式"],
                    "属性值": [file_name, file_size, "嵌入式iframe预览（支持滚动）"]
                }),
                use_container_width=True,
                hide_index=True
            )
        except Exception as e:
            st.warning(f"⚠️ 文件信息获取失败：{str(e)}")
    else:
        # 未选择PDF文件时，显示提示
        st.info("ℹ️ 请在左侧边栏选择PDF来源（上传文件或输入本地路径）以预览报告")

# --------------------------
# 9. 底部说明（补充路径配置与运行注意事项）
# --------------------------
st.divider()
st.markdown("""
### 📌 使用说明
1. **数据来源**：Excel文件路径已固定为 `C:\\Users\\张珊\\Desktop\\3\\数字化转型指数汇总_行业信息完整.xlsx`，无需手动修改；
2. **交互功能**：鼠标悬停在折线图的任意点上，会自动显示对应年份的指数数值（精确到4位小数）；
3. **查询功能**：
   - 企业查询：支持代码/名称模糊匹配，结果含趋势图与历年数据；
   - 行业查询：支持代码/名称匹配，展示行业平均指数趋势；
   - 多行业对比：可选择3-5个行业，对比指数差异与整体平均水平；
   - PDF预览：支持上传或本地路径加载PDF报告，在线预览无需下载；
4. **异常处理**：若数据显示异常，请检查：
   - Excel文件是否存在于固定路径，且未被占用；
   - Excel文件字段名与代码中“required_columns”完全一致（无错别字）；
   - 安装必要依赖（执行 `pip install streamlit pandas plotly openpyxl`）。

### ⚠️ 注意事项
- 企业名称含特殊字符（如*ST、S深发展A）时，输入需完整匹配；
- 多行业对比建议选择3-5个行业，避免图表过于拥挤；
- PDF预览支持最大100MB文件，超大文件可能导致加载缓慢；
- 若Excel读取失败，可手动打开文件确认是否损坏，或重新保存为.xlsx格式后重试。
""")
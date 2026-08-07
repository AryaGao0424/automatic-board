"""
财务共享中心应付组 - 退单分析自动化看板
文件名: fssc_dashboard.py
作者: 数据架构师
说明: 完全基于内置虚拟数据引擎运行，无需读取任何本地 Excel 文件。
      数据结构严格对应真实业务台账的列名与高频词汇。
"""

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from matplotlib import rcParams
import warnings
import random
from datetime import datetime, timedelta

warnings.filterwarnings("ignore")

# ============================================================
# 【全局设置】页面配置与中文字体乱码修复
# ============================================================
st.set_page_config(
    page_title="财务共享中心 · 应付组退单分析看板",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# 修复 matplotlib 中文乱码：优先使用 SimHei，备选 Microsoft YaHei
def set_chinese_font():
    """配置 matplotlib 中文字体，解决乱码问题"""
    font_candidates = ["SimHei", "Microsoft YaHei", "STHeiti", "WenQuanYi Micro Hei"]
    for font in font_candidates:
        try:
            rcParams["font.family"] = font
            rcParams["axes.unicode_minus"] = False  # 解决负号显示为方块的问题
            # 简单测试字体是否可用
            plt.figure(figsize=(1, 1))
            plt.title("测试")
            plt.close()
            break
        except Exception:
            continue

set_chinese_font()

# ============================================================
# 【数据流层】虚拟数据生成引擎
# ============================================================

def generate_mock_fssc_data(n: int = 500, seed: int = 42) -> pd.DataFrame:
    """
    生成高质量虚拟退单明细数据。
    
    参数:
        n (int): 生成数据条数，默认 500 条
        seed (int): 随机种子，保证结果可复现
    
    返回:
        pd.DataFrame: 包含完整业务列名的退单明细 DataFrame
    
    数据分布逻辑（遵循真实业务规律）:
        - 问题类型：附件/影像资料不全 > 未关联合同 > 业务类型填写错误 > 其他
        - 风险等级：中 > 低 > 高（高风险约占 20%）
        - 涉及公司：各子公司退单量不均匀，头部集中效应明显
    """
    np.random.seed(seed)
    random.seed(seed)

    # ── 时间范围：2026年4月1日 ~ 6月30日 ──
    start_date = datetime(2026, 4, 1)
    end_date = datetime(2026, 6, 30)
    date_range_days = (end_date - start_date).days

    # ── 涉及公司（权重不均匀，模拟头部集中效应）──
    companies = [
        "长江创投", "湖北生态", "长江汽车", "长江证券",
        "湖北能源", "长江传媒", "湖北交投", "长江航运",
        "湖北农投", "长江地产"
    ]
    # 前三家公司退单量明显偏高（业务量大、流程不规范）
    company_weights = [0.20, 0.17, 0.15, 0.10, 0.09, 0.08, 0.07, 0.06, 0.05, 0.03]

    # ── 业务场景 ──
    business_scenarios = [
        "支付物料款", "固定资产采购", "工程款支付", "服务费结算",
        "差旅费报销", "租赁费支付", "咨询费结算", "设备维修费"
    ]
    scenario_weights = [0.25, 0.20, 0.18, 0.12, 0.10, 0.07, 0.05, 0.03]

    # ── 问题类型（三大核心类别）──
    problem_types = ["单据和附件", "预算问题", "合同问题"]
    problem_type_weights = [0.50, 0.25, 0.25]

    # ── 问题简述（细化标签，与问题类型对应）──
    # 每个问题类型下的具体问题描述及其权重
    problem_detail_map = {
        "单据和附件": {
            "labels": [
                "附件/影像资料不全",
                "发票信息与单据不符",
                "缺少验收单或签收单",
                "附件扫描件模糊不清",
                "业务类型填写错误",
            ],
            "weights": [0.40, 0.20, 0.18, 0.12, 0.10],
        },
        "预算问题": {
            "labels": [
                "无预算数据支付",
                "超预算金额申请",
                "预算科目填写错误",
                "预算期间不匹配",
            ],
            "weights": [0.45, 0.30, 0.15, 0.10],
        },
        "合同问题": {
            "labels": [
                "未关联合同",
                "合同金额与付款金额不符",
                "合同已到期未续签",
                "合同主体信息有误",
            ],
            "weights": [0.50, 0.25, 0.15, 0.10],
        },
    }

    # ── 风险等级（高:中:低 ≈ 2:5:3）──
    risk_levels = ["高", "中", "低"]
    risk_weights = [0.20, 0.50, 0.30]

    # ── 处理方式 ──
    handle_methods = ["退回前端补充材料", "人工介入核查", "联系业务部门确认", "暂挂待处理"]
    handle_weights = [0.50, 0.25, 0.15, 0.10]

    # ── 处理结果 ──
    handle_results = ["已补充材料重新提交", "已驳回", "处理中", "待确认"]
    handle_result_weights = [0.45, 0.25, 0.20, 0.10]

    # ── 所属业务组 ──
    business_groups = ["应付A组", "应付B组", "应付C组"]
    group_weights = [0.40, 0.35, 0.25]

    # ── 开始逐行生成数据 ──
    records = []
    for i in range(n):
        # 生成发现日期（在时间范围内随机分布，月末略多）
        day_offset = int(np.random.beta(a=2, b=1.5) * date_range_days)
        found_date = start_date + timedelta(days=day_offset)

        # 选择涉及公司
        company = np.random.choice(companies, p=company_weights)

        # 选择业务场景
        scenario = np.random.choice(business_scenarios, p=scenario_weights)

        # 选择问题类型
        prob_type = np.random.choice(problem_types, p=problem_type_weights)

        # 根据问题类型选择具体问题简述
        detail_info = problem_detail_map[prob_type]
        problem_desc = np.random.choice(
            detail_info["labels"], p=detail_info["weights"]
        )

        # 选择风险等级
        risk = np.random.choice(risk_levels, p=risk_weights)

        # 根据风险等级生成涉及金额（高风险金额更大）
        if risk == "高":
            amount = round(np.random.uniform(50, 500), 2)
        elif risk == "中":
            amount = round(np.random.uniform(10, 100), 2)
        else:
            amount = round(np.random.uniform(1, 30), 2)

        # 生成风险点说明（结合问题简述）
        risk_desc_map = {
            "高": f"金额较大，{problem_desc}，可能导致资金损失或合规风险",
            "中": f"{problem_desc}，需及时补充完善，否则影响付款进度",
            "低": f"{problem_desc}，属常规性问题，补充材料后可正常处理",
        }
        risk_desc = risk_desc_map[risk]

        # 选择处理方式与处理结果
        handle = np.random.choice(handle_methods, p=handle_weights)
        result = np.random.choice(handle_results, p=handle_result_weights)

        # 选择所属业务组
        group = np.random.choice(business_groups, p=group_weights)

        records.append({
            "案例编号": f"AP-2026-{str(i + 1).zfill(4)}",
            "发现日期": found_date.strftime("%Y-%m-%d"),
            "所属业务组": group,
            "涉及公司": company,
            "业务场景": scenario,
            "问题类型": prob_type,
            "问题简述": problem_desc,
            "涉及金额（万元）": amount,
            "风险等级": risk,
            "风险点说明": risk_desc,
            "处理方式": handle,
            "处理结果": result,
        })

    df = pd.DataFrame(records)
    # 将发现日期转为 datetime 类型，方便后续时间筛选
    df["发现日期"] = pd.to_datetime(df["发现日期"])
    df["月份"] = df["发现日期"].dt.month
    df["年份"] = df["发现日期"].dt.year
    return df


# ============================================================
# 【主程序入口】加载数据
# ============================================================

@st.cache_data
def load_data():
    """使用缓存加载虚拟数据，避免每次交互重新生成"""
    return generate_mock_fssc_data(n=500)

df_all = load_data()

# ============================================================
# 【侧边栏】过滤控件
# ============================================================

st.sidebar.markdown("## 🔍 数据筛选")
st.sidebar.markdown("---")

# 年份多选（当前数据只有2026年，预留扩展）
all_years = sorted(df_all["年份"].unique().tolist())
selected_years = st.sidebar.multiselect(
    "📅 选择年份",
    options=all_years,
    default=all_years,
    help="可多选年份进行对比分析"
)

# 月份多选
month_map = {4: "4月", 5: "5月", 6: "6月"}
all_months = sorted(df_all["月份"].unique().tolist())
selected_months = st.sidebar.multiselect(
    "📆 选择月份",
    options=all_months,
    format_func=lambda x: month_map.get(x, f"{x}月"),
    default=all_months,
    help="可多选月份进行趋势对比"
)

st.sidebar.markdown("---")

# 风险等级筛选（复选框形式）
st.sidebar.markdown("⚠️ **风险等级筛选**")
risk_high = st.sidebar.checkbox("🔴 高风险", value=True)
risk_mid = st.sidebar.checkbox("🟡 中风险", value=True)
risk_low = st.sidebar.checkbox("🟢 低风险", value=True)

selected_risks = []
if risk_high:
    selected_risks.append("高")
if risk_mid:
    selected_risks.append("中")
if risk_low:
    selected_risks.append("低")

st.sidebar.markdown("---")
st.sidebar.markdown(
    "<small>📌 数据来源：财务共享中心应付组虚拟数据引擎<br>🕐 数据时段：2026年4月-6月</small>",
    unsafe_allow_html=True
)

# ── 应用筛选条件 ──
if not selected_years:
    selected_years = all_years
if not selected_months:
    selected_months = all_months
if not selected_risks:
    selected_risks = ["高", "中", "低"]

df = df_all[
    (df_all["年份"].isin(selected_years)) &
    (df_all["月份"].isin(selected_months)) &
    (df_all["风险等级"].isin(selected_risks))
].copy()

# ============================================================
# 【页面标题】
# ============================================================

st.markdown(
    """
    <div style='background: linear-gradient(90deg, #1a3a5c 0%, #2e6da4 100%);
                padding: 18px 28px; border-radius: 10px; margin-bottom: 20px;'>
        <h2 style='color: white; margin: 0; font-size: 1.6rem;'>
            📊 财务共享中心 · 应付组退单分析看板
        </h2>
        <p style='color: #b8d4f0; margin: 6px 0 0 0; font-size: 0.9rem;'>
            数据时段：2026年4月 - 6月 &nbsp;|&nbsp; 当前筛选：
            {years}年 &nbsp;{months} &nbsp;|&nbsp; 风险等级：{risks}
        </p>
    </div>
    """.format(
        years="/".join(map(str, selected_years)),
        months=" / ".join([month_map.get(m, f"{m}月") for m in selected_months]),
        risks=" / ".join(selected_risks) if selected_risks else "无",
    ),
    unsafe_allow_html=True,
)

# ============================================================
# 【指标聚合层】计算 KPI
# ============================================================

total_cases = len(df)
high_risk_count = len(df[df["风险等级"] == "高"])
high_risk_ratio = (high_risk_count / total_cases * 100) if total_cases > 0 else 0
total_amount = df["涉及金额（万元）"].sum()

# 最高频退单单位
if total_cases > 0:
    top_company = df["涉及公司"].value_counts().idxmax()
    top_company_count = df["涉及公司"].value_counts().max()
else:
    top_company = "—"
    top_company_count = 0

# 最高频问题简述
if total_cases > 0:
    top_problem = df["问题简述"].value_counts().idxmax()
    top_problem_count = df["问题简述"].value_counts().max()
else:
    top_problem = "—"
    top_problem_count = 0

# ============================================================
# 【顶部 KPI 卡片】4 个核心业务指标
# ============================================================

st.markdown("### 📌 核心业务指标")
kpi1, kpi2, kpi3, kpi4 = st.columns(4)

def kpi_card(col, icon, title, value, sub_text, color="#2e6da4"):
    """渲染单个 KPI 卡片"""
    col.markdown(
        f"""
        <div style='background: #f0f6ff; border-left: 5px solid {color};
                    padding: 16px 20px; border-radius: 8px; min-height: 110px;'>
            <div style='font-size: 1.8rem; margin-bottom: 4px;'>{icon}</div>
            <div style='color: #555; font-size: 0.82rem; margin-bottom: 2px;'>{title}</div>
            <div style='color: {color}; font-size: 1.7rem; font-weight: bold;
                        line-height: 1.2;'>{value}</div>
            <div style='color: #888; font-size: 0.78rem; margin-top: 4px;'>{sub_text}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

kpi_card(kpi1, "📋", "本期累计退单笔数", f"{total_cases:,} 笔",
         f"涉及金额合计 {total_amount:,.1f} 万元", color="#2e6da4")

# 预计浪费审批工时 = 总笔数 × 0.5 小时（每笔退单耗费业务与财务双方沟通与重审成本）
wasted_hours = total_cases * 0.5
kpi_card(kpi2, "⏱️", "预计浪费审批工时",
         f"{wasted_hours:,.1f} 小时",
         "亟待前端规范降本", color="#8e44ad")

kpi_card(kpi3, "🏢", "最高频退单单位",
         top_company,
         f"退单 {top_company_count} 笔，占比 {top_company_count/total_cases*100:.1f}%" if total_cases > 0 else "—",
         color="#e67e22")

kpi_card(kpi4, "🏷️", "最高频退单标签",
         top_problem if isinstance(top_problem, str) and len(top_problem) <= 10 else (top_problem[:9] + "…" if isinstance(top_problem, str) else str(top_problem)),
         f"出现 {top_problem_count} 次，占比 {top_problem_count/total_cases*100:.1f}%" if total_cases > 0 else "—",
         color="#27ae60")

# ============================================================
# 【智能业务洞察】动态自然语言结论模块
# ============================================================

if total_cases > 0:
    # 取 Top 2 退单单位
    company_vc = df["涉及公司"].value_counts()
    top1_company = company_vc.index[0] if len(company_vc) >= 1 else "—"
    top2_company = company_vc.index[1] if len(company_vc) >= 2 else "—"

    # 取 Top 1 痛点标签及其占比
    problem_vc = df["问题简述"].value_counts()
    top1_problem = problem_vc.index[0] if len(problem_vc) >= 1 else "—"
    top1_problem_pct = problem_vc.iloc[0] / total_cases * 100 if len(problem_vc) >= 1 else 0

    # 计算向 Top 2 单位发送规范提示函后预计可挽回的工时
    # 假设 Top 2 单位的退单量之和 × 0.5 小时即为可挽回工时
    top2_count = company_vc.iloc[0] + (company_vc.iloc[1] if len(company_vc) >= 2 else 0)
    recoverable_hours = top2_count * 0.5

    insight_text = (
        f"💡 **智能业务洞察：** 当前筛选条件下，退单主要集中在 "
        f"**【{top1_company}】** 与 **【{top2_company}】**，"
        f"核心痛点为 **【{top1_problem}】**（占比 {top1_problem_pct:.1f}%）。"
        f"建议本周由应付组向此两家单位定向发送前端规范提示函，"
        f"预计可挽回约 **{recoverable_hours:.1f} 小时** 的无价值沟通工时。"
    )
    st.warning(insight_text)
else:
    st.info("💡 **智能业务洞察：** 暂无数据支撑洞察，请调整侧边栏筛选条件后重试。")

st.markdown("<br>", unsafe_allow_html=True)

# ============================================================
# 【中部图表区】图表 1 + 图表 2
# ============================================================

st.markdown("### 📈 退单趋势与结构分析")
chart_col1, chart_col2 = st.columns([1.2, 1], gap="large")

# ── 图表 1：三大核心退单原因月度趋势分组柱状图 ──
with chart_col1:
    st.markdown("#### 📅 三大核心退单原因 · 月度趋势")

    # 定义三大核心原因
    core_reasons = ["附件/影像资料不全", "未关联合同", "业务类型填写错误"]
    month_labels = [month_map.get(m, f"{m}月") for m in sorted(selected_months)]

    # 按月份和问题简述聚合
    df_trend = (
        df[df["问题简述"].isin(core_reasons)]
        .groupby(["月份", "问题简述"])
        .size()
        .reset_index(name="退单笔数")
    )

    # 构建绘图数据
    months_sorted = sorted(selected_months)
    x = np.arange(len(months_sorted))
    bar_width = 0.25
    colors_bar = ["#2e6da4", "#e74c3c", "#f39c12"]

    fig1, ax1 = plt.subplots(figsize=(7, 4.2))
    fig1.patch.set_facecolor("#fafcff")
    ax1.set_facecolor("#fafcff")

    for idx, (reason, color) in enumerate(zip(core_reasons, colors_bar)):
        counts = []
        for m in months_sorted:
            val = df_trend[
                (df_trend["月份"] == m) & (df_trend["问题简述"] == reason)
            ]["退单笔数"].sum()
            counts.append(val)
        bars = ax1.bar(
            x + idx * bar_width,
            counts,
            width=bar_width,
            label=reason,
            color=color,
            alpha=0.85,
            edgecolor="white",
            linewidth=0.8,
        )
        # 在柱顶标注数值
        for bar, cnt in zip(bars, counts):
            if cnt > 0:
                ax1.text(
                    bar.get_x() + bar.get_width() / 2,
                    bar.get_height() + 0.5,
                    str(cnt),
                    ha="center", va="bottom",
                    fontsize=8, color="#333"
                )

    ax1.set_xticks(x + bar_width)
    ax1.set_xticklabels(month_labels, fontsize=10)
    ax1.set_ylabel("退单笔数", fontsize=10)
    ax1.set_xlabel("月份", fontsize=10)
    ax1.legend(
        loc="upper right", fontsize=8,
        framealpha=0.7, edgecolor="#ccc"
    )
    ax1.yaxis.set_major_locator(mticker.MaxNLocator(integer=True))
    ax1.spines["top"].set_visible(False)
    ax1.spines["right"].set_visible(False)
    ax1.grid(axis="y", linestyle="--", alpha=0.4)
    plt.tight_layout()
    st.pyplot(fig1)
    plt.close(fig1)

# ── 图表 2：退单痛点结构占比横向条形图 ──
with chart_col2:
    st.markdown("#### 🏷️ 退单痛点 · 问题简述结构占比")

    # 统计各问题简述的频次，取 Top 8
    problem_counts = df["问题简述"].value_counts().head(8)

    if len(problem_counts) > 0:
        labels = problem_counts.index.tolist()[::-1]   # 倒序使最高频在顶部
        raw_vals = problem_counts.values.tolist()[::-1]
        values = [int(str(v).split(".")[0]) if "." in str(v) else int(str(v)) for v in raw_vals]
        total_v = float(sum(values))

        # 颜色渐变：最高频用深色，依次变浅
        cmap = plt.get_cmap("Blues")
        bar_colors = cmap(np.linspace(0.35, 0.85, len(labels)))

        fig2, ax2 = plt.subplots(figsize=(6.5, 4.2))
        fig2.patch.set_facecolor("#fafcff")
        ax2.set_facecolor("#fafcff")

        bars2 = ax2.barh(
            labels, values,
            color=bar_colors,
            edgecolor="white",
            linewidth=0.8,
            height=0.6,
        )

        # 在条形末端标注数值与占比
        for bar, val in zip(bars2, values):
            pct = val / total_v * 100 if total_v > 0 else 0
            ax2.text(
                bar.get_width() + 0.5,
                bar.get_y() + bar.get_height() / 2,
                f"{val} ({pct:.1f}%)",
                va="center", ha="left",
                fontsize=8, color="#333"
            )

        ax2.set_xlabel("退单笔数", fontsize=10)
        ax2.xaxis.set_major_locator(mticker.MaxNLocator(integer=True))
        ax2.spines["top"].set_visible(False)
        ax2.spines["right"].set_visible(False)
        ax2.grid(axis="x", linestyle="--", alpha=0.4)
        # 为最高频条形添加高亮标注
        if values:
            bars2[-1].set_edgecolor("#c0392b")
            bars2[-1].set_linewidth(2)
        plt.tight_layout()
        st.pyplot(fig2)
        plt.close(fig2)
    else:
        st.info("当前筛选条件下无数据，请调整侧边栏筛选项。")

st.markdown("<br>", unsafe_allow_html=True)

# ============================================================
# 【补充图表】各子公司退单量排名（可选展开）
# ============================================================

with st.expander("📊 展开查看：各子公司退单量排名（双轴帕累托图）", expanded=False):

    # ── 任务1：数据预处理 —— 模拟总单量并计算退单率 ──
    company_counts = df["涉及公司"].value_counts().reset_index()
    company_counts.columns = ["委托单位", "退单笔数"]

    # 固定随机种子，保证每次刷新退单率稳定（不随筛选条件漂移）
    rng = np.random.default_rng(seed=99)
    # 各公司退单率在 0.05 ~ 0.35 之间随机模拟
    simulated_rates = rng.uniform(0.05, 0.35, size=len(company_counts))
    # 由退单笔数 / 退单率 反推总单量（取整）
    company_counts["总单量"] = (
        company_counts["退单笔数"] / simulated_rates
    ).astype(int)
    # 真实退单率 = 退单笔数 / 总单量
    company_counts["退单率"] = company_counts["退单笔数"] / company_counts["总单量"]
    # 按退单笔数从高到低排序
    company_counts = company_counts.sort_values("退单笔数", ascending=False).reset_index(drop=True)

    # ── 任务3：结论导向的动态标题 ──
    top_volume_unit = company_counts.iloc[0]["委托单位"] if len(company_counts) > 0 else "—"
    top_rate_unit = company_counts.loc[company_counts["退单率"].idxmax(), "委托单位"] if len(company_counts) > 0 else "—"
    st.markdown(
        f"**🚨 黑榜预警：【{top_volume_unit}】体量居首，"
        f"【{top_rate_unit}】退单率严重超标**"
    )

    # ── 任务2：绘制双坐标轴组合图 ──
    fig3, ax3 = plt.subplots(figsize=(11, 4.2))
    fig3.patch.set_facecolor("#fafcff")
    ax3.set_facecolor("#fafcff")

    x_pos = np.arange(len(company_counts))
    companies_list = company_counts["委托单位"].tolist()
    counts_list = company_counts["退单笔数"].tolist()
    rates_list = company_counts["退单率"].tolist()

    # ── 颜色分配：前3名高亮，支持并列 ──
    sorted_unique_vals = sorted(set(counts_list), reverse=True)
    rank1_val = sorted_unique_vals[0] if len(sorted_unique_vals) >= 1 else None
    rank2_val = sorted_unique_vals[1] if len(sorted_unique_vals) >= 2 else None
    rank3_val = sorted_unique_vals[2] if len(sorted_unique_vals) >= 3 else None

    def get_bar_color_v2(val):
        if val == rank1_val:
            return "#C0392B"
        elif val == rank2_val:
            return "#D35400"
        elif val == rank3_val:
            return "#E67E22"
        else:
            return "#34495E"

    bar_colors3 = [get_bar_color_v2(v) for v in counts_list]

    # 主轴：退单笔数柱状图
    bars3 = ax3.bar(
        x_pos, counts_list,
        color=bar_colors3,
        alpha=0.82,
        edgecolor="white",
        linewidth=0.8,
        width=0.55,
        zorder=2,
    )
    # 柱顶标注退单笔数
    for bar, cnt in zip(bars3, counts_list):
        ax3.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_height() + 0.3,
            str(cnt),
            ha="center", va="bottom",
            fontsize=8.5, color="#333", fontweight="bold"
        )

    ax3.set_ylabel("退单笔数", fontsize=10, color="#34495E")
    ax3.set_xticks(x_pos)
    ax3.set_xticklabels(companies_list, rotation=15, ha="right", fontsize=9)
    ax3.spines["top"].set_visible(False)
    ax3.spines["right"].set_visible(False)
    ax3.grid(axis="y", linestyle="--", alpha=0.35, zorder=0)
    ax3.yaxis.set_major_locator(mticker.MaxNLocator(integer=True))

    # 次轴：退单率折线图
    ax3_r = ax3.twinx()
    ax3_r.plot(
        x_pos, rates_list,
        color="#8E44AD",
        linewidth=2.2,
        marker="o",
        markersize=6,
        markerfacecolor="white",
        markeredgewidth=2,
        markeredgecolor="#8E44AD",
        zorder=3,
        label="退单率",
    )
    # 次轴标签格式化为百分比
    ax3_r.yaxis.set_major_formatter(mticker.FuncFormatter(lambda y, _: f"{y:.0%}"))
    ax3_r.set_ylabel("退单率", fontsize=10, color="#8E44AD")
    ax3_r.tick_params(axis="y", colors="#8E44AD")
    ax3_r.spines["top"].set_visible(False)
    ax3_r.spines["left"].set_visible(False)

    # 管理红线：y=0.15（15%）
    ax3_r.axhline(
        y=0.15,
        color="#E74C3C",
        linestyle="--",
        linewidth=1.5,
        alpha=0.85,
        zorder=4,
    )
    ax3_r.text(
        len(companies_list) - 0.5,
        0.155,
        "管理红线 (15%)",
        color="#E74C3C",
        fontsize=8,
        ha="right",
        va="bottom",
    )

    plt.tight_layout()
    st.pyplot(fig3)
    plt.close(fig3)

# ============================================================
# 【Top 3 重点监管单位体检报告】专项下钻分析模块
# ============================================================

st.markdown("---")
st.subheader("🏥 Top 3 重点监管单位体检报告 (专项下钻)")

if total_cases > 0:
    # ── 提取退单笔数前 3 名的委托单位 ──
    # 复用双轴图中已生成的 company_counts（含总单量与退单率）
    # 若 expander 未展开则 company_counts 可能未定义，此处重新计算以保证独立性
    _cc = df["涉及公司"].value_counts().reset_index()
    _cc.columns = ["委托单位", "退单笔数"]
    _rng = np.random.default_rng(seed=99)
    _rates = _rng.uniform(0.05, 0.35, size=len(_cc))
    _cc["总单量"] = (_cc["退单笔数"] / _rates).astype(int)
    _cc["退单率"] = _cc["退单笔数"] / _cc["总单量"]
    _cc = _cc.sort_values("退单笔数", ascending=False).reset_index(drop=True)

    top3_units = _cc.head(3)["委托单位"].tolist()

    # 莫兰迪/商务色系调色板（深蓝、灰蓝、钢蓝、灰绿、浅灰蓝、暖灰）
    MORANDI_COLORS = [
        "#4A6FA5", "#6B8EB5", "#8DAFC8", "#A8C4D4",
        "#B8CDD6", "#C9D8DE", "#D8E4E8", "#E5EDEF",
    ]

    # ── 用于收集各公司 Top 1 痛点，供督办令使用 ──
    top3_top1_reasons = []  # [(unit_name, top1_reason), ...]

    # 强制指定 4 个切片颜色：深红、深蓝、浅蓝、浅灰
    DONUT_COLORS = ["#C0392B", "#2980B9", "#7FB3D5", "#E5E7E9"]

    col_cards = st.columns(3)

    for idx, unit_name in enumerate(top3_units):
        with col_cards[idx]:
            # ── 1. 单位名称 + 综合评级标签 ──
            unit_row = _cc[_cc["委托单位"] == unit_name].iloc[0]
            unit_rate = unit_row["退单率"]

            if unit_rate > 0.20:
                rating_label = "🔴 极高危 (建议约谈)"
            elif unit_rate >= 0.10:
                rating_label = "🟡 亚健康 (需重点关注)"
            else:
                rating_label = "🟢 运行平稳"

            st.markdown(
                f"**{unit_name}** &nbsp; {rating_label}  \n"
                f"<small style='color:#888;'>退单率：{unit_rate:.1%} | "
                f"退单笔数：{int(unit_row['退单笔数'])} 笔</small>",
                unsafe_allow_html=True,
            )

            # ── 2. 专属痛点环形图（Top 3 + 其他聚合）──
            unit_df = df[df["涉及公司"] == unit_name]
            tag_counts = unit_df["问题简述"].value_counts()

            if len(tag_counts) > 0:
                # ── Top 3 + 其他 聚合逻辑 ──
                if len(tag_counts) > 3:
                    top3_tags = tag_counts.iloc[:3]
                    others_sum = tag_counts.iloc[3:].sum()
                    labels_donut = top3_tags.index.tolist() + ["其他 (Others)"]
                    sizes_donut = top3_tags.values.tolist() + [int(others_sum)]
                else:
                    labels_donut = tag_counts.index.tolist()
                    sizes_donut = tag_counts.values.tolist()

                n_slices = len(labels_donut)

                # 强制使用指定配色（最多 4 个切片）
                slice_colors = DONUT_COLORS[:n_slices]

                # 最大占比项略微 pull out 突出
                explode = [0.06 if i == 0 else 0 for i in range(n_slices)]

                fig_d, ax_d = plt.subplots(figsize=(3, 3))
                fig_d.patch.set_facecolor("none")
                ax_d.set_facecolor("none")

                pie_result = ax_d.pie(
                    sizes_donut,
                    labels=None,
                    colors=slice_colors,
                    explode=explode,
                    autopct=lambda pct: f"{pct:.1f}%" if pct >= 6 else "",
                    pctdistance=0.75,
                    startangle=90,
                    wedgeprops={"linewidth": 1.2, "edgecolor": "white"},
                )
                # autopct 存在时返回 3 元组，否则返回 2 元组，统一兼容处理
                autotexts = pie_result[2] if len(pie_result) == 3 else []

                # 设置百分比字体样式
                for at in autotexts:
                    at.set_fontsize(8)
                    at.set_color("white")
                    at.set_fontweight("bold")

                # 叠加白色圆圈实现环形效果（中心保持纯白，不添加任何文本）
                from matplotlib.patches import Circle as MplCircle
                centre_circle = MplCircle((0, 0), 0.52, fc="white", linewidth=0)
                ax_d.add_patch(centre_circle)

                ax_d.axis("equal")

                # 极简图例：截断超过 10 字的标签，渲染在图表正下方
                legend_labels = [
                    lbl[:10] + "..." if len(lbl) > 10 else lbl
                    for lbl in labels_donut
                ]
                wedges = pie_result[0]
                ax_d.legend(
                    wedges,
                    legend_labels,
                    loc="upper center",
                    bbox_to_anchor=(0.5, -0.05),
                    ncol=1,
                    frameon=False,
                    fontsize=9,
                    handlelength=1.5,
                )

                plt.tight_layout(pad=0.3)
                st.pyplot(fig_d)
                plt.close(fig_d)

                # ── 收集 Top 1 痛点供督办令使用 ──
                top3_top1_reasons.append((unit_name, labels_donut[0]))

            else:
                st.info("该单位在当前筛选条件下暂无退单记录。")
                top3_top1_reasons.append((unit_name, "未知痛点"))

    # ============================================================
    # 【全局动态督办令】横跨全宽，位于 st.columns(3) 完全结束之后
    # ============================================================
    if len(top3_top1_reasons) >= 1:
        # 提取各公司名称与 Top 1 痛点
        _names = [item[0] for item in top3_top1_reasons]
        _reasons = [item[1] for item in top3_top1_reasons]

        # 补齐至 3 个（防止数据不足 3 家公司时报错）
        while len(_names) < 3:
            _names.append("—")
        while len(_reasons) < 3:
            _reasons.append("—")

        top_company_1, top_company_2, top_company_3 = _names[0], _names[1], _names[2]
        top_reason_1, top_reason_2, top_reason_3 = _reasons[0], _reasons[1], _reasons[2]

        # 动态文案：处理痛点重名情况
        if top_reason_2 == top_reason_3:
            # 两家公司痛点相同，合并表述
            if top_reason_1 == top_reason_2:
                # 三家痛点全部相同
                order_text = (
                    f"经系统穿透，【{top_company_1}】、【{top_company_2}】与【{top_company_3}】"
                    f"的核心合规卡点高度一致，均集中在'{top_reason_1}'。"
                    f"建议应付组将上述痛点列入本月严控指标，"
                    f"针对性下发《前端报账规范提示函》，并抄送对应单位财务分管领导。"
                )
            else:
                order_text = (
                    f"经系统穿透，【{top_company_1}】的核心合规卡点为'{top_reason_1}'，"
                    f"【{top_company_2}】与【{top_company_3}】的高频违规点均集中在'{top_reason_2}'。"
                    f"建议应付组将上述痛点列入本月严控指标，"
                    f"针对性下发《前端报账规范提示函》，并抄送对应单位财务分管领导。"
                )
        elif top_reason_1 == top_reason_2:
            # 前两家痛点相同
            order_text = (
                f"经系统穿透，【{top_company_1}】与【{top_company_2}】的核心合规卡点均为'{top_reason_1}'，"
                f"【{top_company_3}】的高频违规点集中在'{top_reason_3}'。"
                f"建议应付组将上述痛点列入本月严控指标，"
                f"针对性下发《前端报账规范提示函》，并抄送对应单位财务分管领导。"
            )
        elif top_reason_1 == top_reason_3:
            # 第一和第三家痛点相同
            order_text = (
                f"经系统穿透，【{top_company_1}】与【{top_company_3}】的核心合规卡点均为'{top_reason_1}'，"
                f"【{top_company_2}】的高频违规点集中在'{top_reason_2}'。"
                f"建议应付组将上述痛点列入本月严控指标，"
                f"针对性下发《前端报账规范提示函》，并抄送对应单位财务分管领导。"
            )
        else:
            # 三家痛点各不相同
            order_text = (
                f"经系统穿透，【{top_company_1}】的核心合规卡点为'{top_reason_1}'，"
                f"【{top_company_2}】与【{top_company_3}】的高频违规点集中在"
                f"'{top_reason_2}'与'{top_reason_3}'。"
                f"建议应付组将上述痛点列入本月严控指标，"
                f"针对性下发《前端报账规范提示函》，并抄送对应单位财务分管领导。"
            )

        st.error(f"🚨 **总中心穿透督办令**\n\n{order_text}")

else:
    st.info("暂无数据，请调整侧边栏筛选条件后重试。")

st.markdown("<br>", unsafe_allow_html=True)

# ============================================================
# 【BIP系统优化与攻坚进度追踪】模块
# ============================================================

def generate_mock_system_issues() -> pd.DataFrame:
    """
    生成高度拟真的 Yonyou BIP 系统优化需求明细数据（约30条）。
    数据贴合大型国企财务共享中心真实业务场景。
    """
    tasks = [
        "凭证摘要取值规则优化：自动带出业务单据摘要字段",
        "预算数据占用异常报错：提交付款申请时提示预算余额不足但实际有余额",
        "发票专票税率无法自动识别：17%税率发票OCR识别后税率字段为空",
        "流水生单无法关联预付单：付款流水与预付款单据关联逻辑缺失",
        "供应商银行账号变更后付款单自动带出旧账号问题",
        "合同付款计划与实际付款进度不同步，导致超付预警误报",
        "费用报销单审批流程中，部门负责人节点无法委托代审",
        "固定资产卡片折旧计算异常：当月新增资产当月计提折旧",
        "应付账款对账单批量导出功能缺失，需逐条手工下载",
        "发票池中已认证发票无法与多张付款单进行拆分匹配",
        "预算调整申请单审批通过后，预算系统余额未实时刷新",
        "工程款支付申请中，工程进度节点附件上传大小限制过严（仅5MB）",
        "跨公司内部往来结算单据，对方公司确认环节无消息推送通知",
        "付款申请单提交后状态长时间停留'待审批'，审批人未收到待办",
        "月末批量生成凭证时，系统超时报错，无法完成批量过账",
        "供应商信息维护界面，统一社会信用代码校验规则未启用",
        "资金计划与付款申请关联时，计划编号搜索框不支持模糊查询",
        "电子发票归档后，在凭证附件中无法直接预览PDF，需下载后查看",
        "三单匹配（采购订单/收货单/发票）差异超容差时，系统无差异明细提示",
        "付款单据打印模板中，大写金额字段在金额超千万时显示格式错误",
        "预付款核销时，系统不允许跨期核销，导致跨年预付款无法正常处理",
        "费用报销单中，差旅费明细行超过20行时，页面加载卡顿严重",
        "银行回单自动匹配功能：同一天同金额多笔流水无法精准匹配",
        "合同台账中，合同到期预警邮件推送功能未上线",
        "付款申请单驳回后，原单据附件丢失，需重新上传",
        "增值税进项税额转出业务，系统无专用单据支持，需手工调整凭证",
        "应付账款账龄分析报表，账龄区间设置不支持自定义",
        "资金支付指令下发后，银企直联回执状态未自动回写至付款单",
        "期末结账后，已关闭期间的凭证仍可被误操作修改",
        "BIP移动端审批APP，附件图片旋转方向异常，影响审核判断",
    ]

    priorities_pool = ["🔥 紧急重要", "不紧急&重要", "紧急&不重要"]
    priority_weights = [0.55, 0.30, 0.15]

    statuses_pool = ["已完成", "已解决", "正在处理", "正在沟通解决中", "无法完成"]
    status_weights = [0.25, 0.20, 0.25, 0.20, 0.10]

    owners_pool = ["王磊（实施顾问）", "张敏（财务IT）", "李强（用友技术）", "陈静（应付组长）", "刘洋（系统管理员）"]

    remarks_map = {
        "已完成":       ["已于上线补丁包中修复，测试通过", "需求已实现，已通知业务端验收", "配置调整完毕，运行正常"],
        "已解决":       ["临时方案已上线，待正式版本迭代", "通过参数配置解决，已关闭工单", "用友远程协助处理完毕"],
        "正在处理":     ["用友研发已排期，预计下月版本修复", "内部测试环境已复现，等待补丁", "已提交用友工单，跟进中"],
        "正在沟通解决中": ["与用友实施顾问持续沟通方案中", "涉及底层逻辑改造，需评估工作量", "已升级至用友产品部，等待回复"],
        "无法完成":     ["受限于标准产品架构，暂无解决方案", "用友评估后确认为产品设计逻辑，不予修改", "需二次开发，成本过高暂搁置"],
    }

    random.seed(2026)
    np.random.seed(2026)

    records = []
    for i, task in enumerate(tasks):
        priority = np.random.choice(priorities_pool, p=priority_weights)
        status = np.random.choice(statuses_pool, p=status_weights)
        owner = random.choice(owners_pool)
        remark = random.choice(remarks_map[status])
        records.append({
            "任务内容": task,
            "优先级": priority,
            "进度状态": status,
            "责任人": owner,
            "未完成原因/跟进备注": remark,
        })

    return pd.DataFrame(records)


# ── 加载系统优化数据（缓存）──
@st.cache_data
def load_system_issues():
    return generate_mock_system_issues()

df_issues = load_system_issues()

# ── 分割线与模块标题 ──
st.markdown("---")
st.subheader("🛠️ Yonyou BIP 系统优化与攻坚追踪台账")

# ── 顶部 KPI 卡片（4列）──
bip_kpi1, bip_kpi2, bip_kpi3, bip_kpi4 = st.columns(4)

total_issues = 67
urgent_important_count = 55
urgent_important_pct = 82
closed_count = 44
closed_pct = 65
inprogress_count = 23

with bip_kpi1:
    st.markdown(
        f"""
        <div style='background:#f0f6ff; border-left:5px solid #2e6da4;
                    padding:16px 20px; border-radius:8px; min-height:110px;'>
            <div style='font-size:1.8rem; margin-bottom:4px;'>📌</div>
            <div style='color:#555; font-size:0.82rem; margin-bottom:2px;'>累计提出系统需求</div>
            <div style='color:#2e6da4; font-size:1.7rem; font-weight:bold; line-height:1.2;'>{total_issues} 项</div>
            <div style='color:#888; font-size:0.78rem; margin-top:4px;'>覆盖全流程核心痛点</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with bip_kpi2:
    st.markdown(
        f"""
        <div style='background:#fff8f0; border-left:5px solid #e67e22;
                    padding:16px 20px; border-radius:8px; min-height:110px;'>
            <div style='font-size:1.8rem; margin-bottom:4px;'>🔥</div>
            <div style='color:#555; font-size:0.82rem; margin-bottom:2px;'>紧急且重要占比</div>
            <div style='color:#e67e22; font-size:1.7rem; font-weight:bold; line-height:1.2;'>{urgent_important_pct}%</div>
            <div style='color:#888; font-size:0.78rem; margin-top:4px;'>共 {urgent_important_count} 项，亟待攻坚</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with bip_kpi3:
    st.markdown(
        f"""
        <div style='background:#f0fff4; border-left:5px solid #27ae60;
                    padding:16px 20px; border-radius:8px; min-height:110px;'>
            <div style='font-size:1.8rem; margin-bottom:4px;'>✅</div>
            <div style='color:#555; font-size:0.82rem; margin-bottom:2px;'>已闭环解决/完成</div>
            <div style='color:#27ae60; font-size:1.7rem; font-weight:bold; line-height:1.2;'>{closed_count} 项</div>
            <div style='color:#888; font-size:0.78rem; margin-top:4px;'>整体进度约 {closed_pct}%</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with bip_kpi4:
    st.markdown(
        f"""
        <div style='background:#fdf0ff; border-left:5px solid #8e44ad;
                    padding:16px 20px; border-radius:8px; min-height:110px;'>
            <div style='font-size:1.8rem; margin-bottom:4px;'>⚙️</div>
            <div style='color:#555; font-size:0.82rem; margin-bottom:2px;'>当前攻坚/跟进中</div>
            <div style='color:#8e44ad; font-size:1.7rem; font-weight:bold; line-height:1.2;'>{inprogress_count} 项</div>
            <div style='color:#888; font-size:0.78rem; margin-top:4px;'>持续推进，紧盯落地</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

st.markdown("<br>", unsafe_allow_html=True)

# ── 进度状态多选筛选器 ──
all_statuses = df_issues["进度状态"].unique().tolist()
selected_statuses = st.multiselect(
    "🔎 按进度状态筛选（可多选）",
    options=all_statuses,
    default=all_statuses,
    help="选择一个或多个进度状态，过滤下方明细表",
    key="bip_status_filter",
)

if not selected_statuses:
    selected_statuses = all_statuses

df_issues_filtered = df_issues[df_issues["进度状态"].isin(selected_statuses)].copy()

# ── 交互式数据表（带 column_config 美化）──
st.dataframe(
    df_issues_filtered,
    use_container_width=True,
    height=520,
    column_config={
        "任务内容": st.column_config.TextColumn(
            label="📝 任务内容 / 业务痛点",
            width="large",
            help="来源于财务共享中心日常操作中发现的系统缺陷与优化需求",
        ),
        "优先级": st.column_config.TextColumn(
            label="⚡ 优先级",
            width="medium",
            help="🔥 紧急重要 = 需立即推动解决；不紧急&重要 = 纳入迭代计划；紧急&不重要 = 临时处理",
        ),
        "进度状态": st.column_config.TextColumn(
            label="📊 进度状态",
            width="medium",
            help="当前该需求/缺陷的处理进展",
        ),
        "责任人": st.column_config.TextColumn(
            label="👤 责任人",
            width="small",
        ),
        "未完成原因/跟进备注": st.column_config.TextColumn(
            label="💬 未完成原因 / 跟进备注",
            width="large",
            help="记录当前阻塞原因或最新跟进进展",
        ),
    },
    hide_index=True,
)

st.markdown(
    f"<small style='color:#888;'>共展示 {len(df_issues_filtered)} 条 / 台账合计 {len(df_issues)} 条系统优化需求</small>",
    unsafe_allow_html=True,
)

st.markdown("<br>", unsafe_allow_html=True)

# ============================================================
# 【底部数据表】退单明细一维表
# ============================================================

st.markdown("### 📋 退单明细底层数据（规范样式预览）")

# 格式化展示列（隐藏辅助列 月份/年份）
display_cols = [
    "案例编号", "发现日期", "所属业务组", "涉及公司",
    "业务场景", "问题类型", "问题简述",
    "涉及金额（万元）", "风险等级", "风险点说明",
    "处理方式", "处理结果"
]

# 风险等级颜色映射（用于高亮显示）
def highlight_risk(val):
    """根据风险等级返回对应的背景色样式"""
    color_map = {"高": "#fde8e8", "中": "#fff8e1", "低": "#e8f5e9"}
    return f"background-color: {color_map.get(val, 'white')}"

df_display = df[display_cols].copy()
df_display["发现日期"] = df_display["发现日期"].dt.strftime("%Y-%m-%d")

# 展示行数控制
show_rows = st.slider(
    "📌 展示行数",
    min_value=10, max_value=min(200, len(df_display)),
    value=min(50, len(df_display)), step=10,
    help="拖动滑块控制底部数据表展示的行数"
)

st.dataframe(
    df_display.head(show_rows).style.map(
        highlight_risk, subset=["风险等级"]
    ).format({"涉及金额（万元）": "{:.2f}"}),
    use_container_width=True,
    height=420,
)

st.markdown(
    f"<small style='color:#888;'>共展示 {show_rows} 条 / 筛选后合计 {total_cases} 条退单记录</small>",
    unsafe_allow_html=True,
)

# ============================================================
# 【页脚】
# ============================================================
st.markdown("---")
st.markdown(
    "<div style='text-align:center; color:#aaa; font-size:0.8rem;'>"
    "财务共享中心应付组 · 退单分析自动化看板 &nbsp;|&nbsp; "
    "数据为虚拟生成，仅供演示 &nbsp;|&nbsp; Powered by Streamlit"
    "</div>",
    unsafe_allow_html=True,
)

"""
财务共享中心应付组 - 退单分析自动化看板
文件名: fssc_dashboard.py
作者: 数据架构师
说明: 基于真实业务数据驱动，通过侧边栏上传两个Excel文件（主表+金额表）。
      数据来源：财务共享中心BIP系统导出数据（2026年1月-6月）
"""

import re
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from matplotlib import rcParams
from matplotlib.patches import Circle as MplCircle, Rectangle as MplRectangle
import warnings
import random

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

def set_chinese_font():
    font_candidates = ["SimHei", "Microsoft YaHei", "STHeiti", "WenQuanYi Micro Hei"]
    for font in font_candidates:
        try:
            rcParams["font.family"] = font
            rcParams["axes.unicode_minus"] = False
            plt.figure(figsize=(1, 1))
            plt.title("测试")
            plt.close()
            break
        except Exception:
            continue

set_chinese_font()

# ============================================================
# 【工具函数】
# ============================================================

def parse_duration_seconds(s):
    """解析 'X分X秒' 格式为秒数；'--' 视为0秒"""
    if pd.isna(s) or str(s).strip() == "":
        return None
    s = str(s).strip()
    if s == "--":
        return 0
    m = re.match(r"(?:(\d+)分)?(\d+)秒", s)
    if m:
        mins = int(m.group(1)) if m.group(1) else 0
        secs = int(m.group(2))
        return mins * 60 + secs
    return None

def is_rejected_flag(x):
    """判断是否曾被退回（驳出共享次数 >= 1）"""
    s = str(x).strip()
    if s in ("--", "nan", "0", ""):
        return False
    try:
        return int(float(s)) >= 1
    except Exception:
        return False

@st.cache_data
def load_and_merge(main_bytes, amount_bytes):
    """读取并合并主表与金额表"""
    df_main = pd.read_excel(main_bytes, sheet_name=0)
    df_amount = pd.read_excel(amount_bytes, sheet_name=0)

    # 金额表：按单据编号分组求和（避免一对多导致行数膨胀）
    df_amt_agg = (
        df_amount.groupby("单据编号", as_index=False)["金额"]
        .sum()
        .rename(columns={"金额": "金额（元）"})
    )

    # left join
    df = df_main.merge(df_amt_agg, on="单据编号", how="left")

    # 金额换算：元 → 万元
    df["金额（万元）"] = df["金额（元）"] / 10000

    # 解析处理时间
    df["处理时间_dt"] = pd.to_datetime(df["处理时间"], errors="coerce")
    df["月份"] = df["处理时间_dt"].dt.month
    df["年份"] = df["处理时间_dt"].dt.year

    # 解析作业时长（使用修正作业时长，'--' 视为0秒）
    df["作业时长_秒"] = df["修正作业时长"].apply(parse_duration_seconds)

    # 标记曾被退回（口径②）
    df["is_rejected"] = df["驳出共享次数"].apply(is_rejected_flag)

    return df

# ============================================================
# 【侧边栏】文件上传 + 过滤控件
# ============================================================

st.sidebar.markdown("## 📁 数据上传")
st.sidebar.markdown("---")

main_file = st.sidebar.file_uploader(
    "上传主表（作业明细查询列表）",
    type=["xlsx"],
    help="请上传：1-6月应付组作业明细查询列表.xlsx",
    key="main_file",
)
amount_file = st.sidebar.file_uploader(
    "上传金额表（作业明细查询-金额）",
    type=["xlsx"],
    help="请上传：1-6月应付组作业明细查询-金额.xlsx",
    key="amount_file",
)
bip_file = st.sidebar.file_uploader(
    "上传BIP系统优化台账",
    type=["xlsx"],
    help="请上传：BIP应付组系统优化点_分类汇总.xlsx（读取【分类明细】sheet）",
    key="bip_file",
)

st.sidebar.markdown("---")

# ── 数据加载 ──
if main_file is None or amount_file is None:
    st.markdown(
        """
        <div style='background: linear-gradient(90deg, #1a3a5c 0%, #2e6da4 100%);
                    padding: 18px 28px; border-radius: 10px; margin-bottom: 20px;'>
            <h2 style='color: white; margin: 0; font-size: 1.6rem;'>
                📊 财务共享中心 · 应付组退单分析看板
            </h2>
            <p style='color: #b8d4f0; margin: 6px 0 0 0; font-size: 0.9rem;'>
                数据时段：2026年1月 - 6月 &nbsp;|&nbsp; 数据来源：财务共享中心BIP系统
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.info(
        "👈 **请在左侧侧边栏上传两个数据文件后，看板将自动加载：**\n\n"
        "1. **主表**：`1-6月应付组作业明细查询列表.xlsx`\n"
        "2. **金额表**：`1-6月应付组作业明细查询-金额.xlsx`\n"
        "3. **系统优化台账**：`BIP应付组系统优化点_分类汇总.xlsx`"
    )
    st.stop()

df_all = load_and_merge(main_file, amount_file)

# ── 侧边栏筛选控件 ──
st.sidebar.markdown("## 🔍 数据筛选")
st.sidebar.markdown("---")

month_map = {1: "1月", 2: "2月", 3: "3月", 4: "4月", 5: "5月", 6: "6月"}
all_months = sorted([int(m) for m in df_all["月份"].dropna().unique().tolist() if 1 <= int(m) <= 12])
selected_months = st.sidebar.multiselect(
    "📆 选择月份",
    options=all_months,
    format_func=lambda x: month_map.get(int(x), f"{int(x)}月"),
    default=all_months,
    help="按处理时间月份筛选",
)

st.sidebar.markdown("---")

all_handlers = sorted(df_all["处理人"].dropna().unique().tolist())
selected_handlers = st.sidebar.multiselect(
    "👤 处理人筛选",
    options=all_handlers,
    default=all_handlers,
    help="选择审核员",
)

all_statuses_work = df_all["作业状态"].dropna().unique().tolist()
selected_work_statuses = st.sidebar.multiselect(
    "📋 作业状态筛选",
    options=all_statuses_work,
    default=all_statuses_work,
    help="审核通过 / 审核驳回",
)

all_trade_types = sorted(df_all["交易类型"].dropna().unique().tolist())
selected_trade_types = st.sidebar.multiselect(
    "🔄 交易类型筛选",
    options=all_trade_types,
    default=all_trade_types,
    help="选择单据交易类型",
)

st.sidebar.markdown("---")
st.sidebar.markdown(
    "<small>📌 数据来源：财务共享中心BIP系统导出<br>🕐 数据时段：2026年1月-6月</small>",
    unsafe_allow_html=True,
)

# ── 应用筛选条件 ──
if not selected_months:
    selected_months = all_months
if not selected_handlers:
    selected_handlers = all_handlers
if not selected_work_statuses:
    selected_work_statuses = all_statuses_work
if not selected_trade_types:
    selected_trade_types = all_trade_types

df = df_all[
    (df_all["月份"].isin(selected_months)) &
    (df_all["处理人"].isin(selected_handlers)) &
    (df_all["作业状态"].isin(selected_work_statuses)) &
    (df_all["交易类型"].isin(selected_trade_types))
].copy()

# ============================================================
# 【页面标题】
# ============================================================

months_str = " / ".join([month_map.get(int(m), f"{int(m)}月") for m in sorted(selected_months)])
st.markdown(
    f"""
    <div style='background: linear-gradient(90deg, #1a3a5c 0%, #2e6da4 100%);
                padding: 18px 28px; border-radius: 10px; margin-bottom: 20px;'>
        <h2 style='color: white; margin: 0; font-size: 1.6rem;'>
            📊 财务共享中心 · 应付组退单分析看板
        </h2>
        <p style='color: #b8d4f0; margin: 6px 0 0 0; font-size: 0.9rem;'>
            数据时段：2026年1月 - 6月 &nbsp;|&nbsp; 当前筛选月份：{months_str}
            &nbsp;|&nbsp; 处理人：{len(selected_handlers)}人 &nbsp;|&nbsp;
            作业状态：{"、".join(selected_work_statuses)}
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# 【指标聚合层】计算 KPI
# ============================================================

# 全量去重后的唯一单据（退单率分母）
df_unique_all = df.drop_duplicates(subset="单据编号", keep="last")
total_unique_docs = len(df_unique_all)

# 口径②：曾被退回的唯一单据（退单率分子）
df_rejected_unique = df[df["is_rejected"]].drop_duplicates(subset="单据编号", keep="first")
rejected_unique_count = len(df_rejected_unique)
rejection_rate = rejected_unique_count / total_unique_docs if total_unique_docs > 0 else 0

# 退单记录（作业状态=审核驳回）的整改耗时
df_rej_records = df[df["作业状态"] == "审核驳回"]
total_reject_seconds = df_rej_records["作业时长_秒"].sum()
total_reject_hours = total_reject_seconds / 3600 if total_reject_seconds and total_reject_seconds > 0 else 0
avg_min_per_case = (total_reject_seconds / len(df_rej_records) / 60) if len(df_rej_records) > 0 and total_reject_seconds else 0

# 最高频退单委托单位（口径②）
if rejected_unique_count > 0:
    top_company_vc = df_rejected_unique["委托单位"].value_counts()
    top_company = top_company_vc.index[0]
    top_company_count = int(top_company_vc.iloc[0])
else:
    top_company = "—"
    top_company_count = 0

# 最高频退单标签
rej_labels = df_rej_records["处理意见标签化"].dropna()
if len(rej_labels) > 0:
    top_label_vc = rej_labels.value_counts()
    top_label = top_label_vc.index[0]
    top_label_count = int(top_label_vc.iloc[0])
    top_label_total = len(rej_labels)
else:
    top_label = "—"
    top_label_count = 0
    top_label_total = 0

# 退单金额最大委托单位（口径②）
if rejected_unique_count > 0:
    unit_amount_kpi = df_rejected_unique.groupby("委托单位")["金额（万元）"].sum()
    unit_amount_kpi = unit_amount_kpi[unit_amount_kpi > 0].sort_values(ascending=False)
    if len(unit_amount_kpi) > 0:
        top_amount_unit = unit_amount_kpi.index[0]
        top_amount_val = unit_amount_kpi.iloc[0]
    else:
        top_amount_unit = "—"
        top_amount_val = 0.0
else:
    top_amount_unit = "—"
    top_amount_val = 0.0

# ============================================================
# 【顶部 KPI 卡片】5 个核心业务指标
# ============================================================

st.markdown("### 📌 核心业务指标")
kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)

def kpi_card(col, icon, title, value, sub_text, color="#2e6da4"):
    col.markdown(
        f"""
        <div style='background: #f0f6ff; border-left: 5px solid {color};
                    padding: 14px 16px; border-radius: 8px; min-height: 120px;'>
            <div style='font-size: 1.6rem; margin-bottom: 4px;'>{icon}</div>
            <div style='color: #555; font-size: 0.78rem; margin-bottom: 2px;'>{title}</div>
            <div style='color: {color}; font-size: 1.45rem; font-weight: bold;
                        line-height: 1.2;'>{value}</div>
            <div style='color: #888; font-size: 0.74rem; margin-top: 4px;'>{sub_text}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

kpi_card(
    kpi1, "📋", "被退回单据数（剔除重复计数）",
    f"{rejected_unique_count:,} 笔",
    f"退单率 {rejection_rate:.1%}（共 {total_unique_docs:,} 笔唯一单据）",
    color="#2e6da4",
)
kpi_card(
    kpi2, "⏱️", "退单整改耗时",
    f"{total_reject_hours:,.1f} 小时",
    f"退单 {len(df_rej_records)} 笔，均 {avg_min_per_case:.1f} 分/笔" if len(df_rej_records) > 0 else "暂无数据",
    color="#8e44ad",
)
kpi_card(
    kpi3, "🏢", "最高频退单委托单位",
    top_company[:12] + "…" if len(top_company) > 12 else top_company,
    f"退单 {top_company_count} 笔，占比 {top_company_count/rejected_unique_count*100:.1f}%" if rejected_unique_count > 0 else "—",
    color="#e67e22",
)
kpi_card(
    kpi4, "🏷️", "最高频退单标签",
    top_label[:10] + "…" if len(top_label) > 10 else top_label,
    f"出现 {top_label_count} 次，占比 {top_label_count/top_label_total*100:.1f}%" if top_label_total > 0 else "—",
    color="#27ae60",
)
kpi_card(
    kpi5, "💰", "退单金额最大委托单位",
    top_amount_unit[:10] + "…" if len(top_amount_unit) > 10 else top_amount_unit,
    f"退单涉及金额 {top_amount_val:,.1f} 万元",
    color="#c0392b",
)

# ============================================================
# 【智能业务洞察】
# ============================================================

if rejected_unique_count > 0:
    company_vc2 = df_rejected_unique["委托单位"].value_counts()
    top1_c = company_vc2.index[0] if len(company_vc2) >= 1 else "—"
    top2_c = company_vc2.index[1] if len(company_vc2) >= 2 else "—"
    label_vc2 = rej_labels.value_counts() if len(rej_labels) > 0 else pd.Series(dtype=int)
    top1_label = label_vc2.index[0] if len(label_vc2) >= 1 else "—"
    top1_label_pct = label_vc2.iloc[0] / len(rej_labels) * 100 if len(label_vc2) >= 1 else 0
    top2_rej_count = int(company_vc2.iloc[0]) + (int(company_vc2.iloc[1]) if len(company_vc2) >= 2 else 0)
    avg_sec = total_reject_seconds / len(df_rej_records) if len(df_rej_records) > 0 and total_reject_seconds else 1800
    recoverable_hours = top2_rej_count * avg_sec / 3600
    insight_text = (
        f"💡 **智能业务洞察：** 当前筛选条件下，退单主要集中在 "
        f"**【{top1_c}】** 与 **【{top2_c}】**，"
        f"核心痛点为 **【{top1_label}】**（占比 {top1_label_pct:.1f}%）。"
        f"如果应付组及时向此两家单位提供定向沟通服务，针对性解决退单痛点，有效降低退单率，"
        f"则预计可节省约 **{recoverable_hours:.1f} 小时** 的退单整改耗时。"
    )
    st.warning(insight_text)
else:
    st.info("💡 **智能业务洞察：** 暂无退单数据，请调整侧边栏筛选条件后重试。")

st.markdown("<br>", unsafe_allow_html=True)

# ============================================================
# 【中部图表区】图表 1 + 图表 2
# ============================================================

st.markdown("### 📈 退单趋势与结构分析")
chart_col1, chart_col2 = st.columns([1.2, 1], gap="large")

with chart_col1:
    st.markdown("#### 📅 三大核心退单原因 · 月度趋势")
    core_reasons = ["附件不完整", "附件信息不一致", "其他"]
    months_sorted = sorted([int(m) for m in selected_months])
    month_labels = [month_map.get(m, f"{m}月") for m in months_sorted]

    df_trend = (
        df_rej_records[df_rej_records["处理意见标签化"].isin(core_reasons)]
        .groupby(["月份", "处理意见标签化"])
        .size()
        .reset_index(name="退单笔数")
    )

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
                (df_trend["月份"] == m) & (df_trend["处理意见标签化"] == reason)
            ]["退单笔数"].sum()
            counts.append(int(val))
        bars = ax1.bar(
            x + idx * bar_width, counts,
            width=bar_width, label=reason,
            color=color, alpha=0.85, edgecolor="white", linewidth=0.8,
        )
        for bar, cnt in zip(bars, counts):
            if cnt > 0:
                ax1.text(
                    bar.get_x() + bar.get_width() / 2,
                    bar.get_height() + 0.3,
                    str(cnt), ha="center", va="bottom", fontsize=8, color="#333",
                )

    ax1.set_xticks(x + bar_width)
    ax1.set_xticklabels(month_labels, fontsize=10)
    ax1.set_ylabel("退单笔数", fontsize=10)
    ax1.set_xlabel("月份（按处理时间）", fontsize=10)
    ax1.legend(loc="upper right", fontsize=8, framealpha=0.7, edgecolor="#ccc")
    ax1.yaxis.set_major_locator(mticker.MaxNLocator(integer=True))
    ax1.spines["top"].set_visible(False)
    ax1.spines["right"].set_visible(False)
    ax1.grid(axis="y", linestyle="--", alpha=0.4)
    plt.tight_layout()
    st.pyplot(fig1)
    plt.close(fig1)

with chart_col2:
    st.markdown("#### 🏷️ 退单痛点 · 处理意见结构占比")
    label_counts = df_rej_records["处理意见标签化"].dropna().value_counts()

    if len(label_counts) > 0:
        # "特殊原因"固定置底（水平条形图最下方），其余按退单数量由多到少从上向下排列
        SPECIAL_LABEL = "特殊原因"
        if SPECIAL_LABEL in label_counts.index:
            special_val = int(label_counts[SPECIAL_LABEL])
            other_counts = label_counts.drop(index=SPECIAL_LABEL).sort_values(ascending=False)
            # 水平条形图：列表顺序 = 从下到上，故先放特殊原因（最下），再放其他（从下到上按数量升序）
            other_vals_rev = list(reversed(other_counts.astype(int).tolist()))
            labels_bar = [SPECIAL_LABEL] + [str(lbl) for lbl in other_counts.index.tolist()[::-1]]
            values_bar = [special_val] + other_vals_rev
        else:
            labels_bar = [str(lbl) for lbl in label_counts.index.tolist()[::-1]]
            values_bar = list(reversed(label_counts.astype(int).tolist()))
        total_v = float(sum(values_bar))
        cmap = plt.get_cmap("Blues")
        bar_colors2 = cmap(np.linspace(0.35, 0.85, len(labels_bar)))

        fig2, ax2 = plt.subplots(figsize=(6.5, 4.2))
        fig2.patch.set_facecolor("#fafcff")
        ax2.set_facecolor("#fafcff")

        bars2 = ax2.barh(labels_bar, values_bar, color=bar_colors2, edgecolor="white", linewidth=0.8, height=0.6)
        for bar, val in zip(bars2, values_bar):
            pct = val / total_v * 100 if total_v > 0 else 0
            ax2.text(bar.get_width() + 0.3, bar.get_y() + bar.get_height() / 2,
                     f"{val} ({pct:.1f}%)", va="center", ha="left", fontsize=8, color="#333")

        ax2.set_xlabel("退单笔数", fontsize=10)
        ax2.xaxis.set_major_locator(mticker.MaxNLocator(integer=True))
        ax2.spines["top"].set_visible(False)
        ax2.spines["right"].set_visible(False)
        ax2.grid(axis="x", linestyle="--", alpha=0.4)
        if values_bar:
            bars2[-1].set_edgecolor("#c0392b")
            bars2[-1].set_linewidth(2)
        plt.tight_layout()
        st.pyplot(fig2)
        plt.close(fig2)
    else:
        st.info("当前筛选条件下无退单数据，请调整侧边栏筛选项。")

st.markdown("<br>", unsafe_allow_html=True)

# ============================================================
# 【补充图表】各委托单位退单量排名（双柱形图 + 折线图）
# ============================================================

st.markdown("#### 📊 各委托单位退单分析（上图：退单笔数 + 退单金额；下图：退单率）")
st.info(
    "**📐 指标计算逻辑说明：**\n\n"
    "- **退单笔数**：曾被退回的唯一单据数，按单据编号去重，以驳出共享次数 ≥ 1 为判断标准。\n"
    "- **退单率**：该委托单位退单笔数 ÷ 该委托单位总单量（唯一单据数），反映各单位的退单频率。\n"
    "- **退单涉及金额（万元）**：退单唯一单据对应的金额合计（元转万元），金额来源于金额表，按单据编号匹配后求和；受单位主营业务类型影响较大，仅作风险参考，同笔数时作为次级排序键。\n"
    "- **排序口径**：上下两图选取同一批【退单笔数由高到低】的前 20 个委托单位；上图按【退单笔数】降序排列（体量维度，同笔数者按退单金额降序），下图对同一批单位按【退单率由高到低】重新排序（质量维度），便于分别聚焦体量风险与质量风险。\n"
    "- **管理红线（15%）**：退单率超过15%需重点关注。"
)



unit_rej = (
    df[df["is_rejected"]]
    .drop_duplicates(subset="单据编号", keep="first")
    .groupby("委托单位")
    .agg(退单笔数=("单据编号", "count"), 退单金额万元=("金额（万元）", "sum"))
    .reset_index()
)
unit_total = (
    df.drop_duplicates(subset="单据编号", keep="last")
    .groupby("委托单位")
    .agg(总单量=("单据编号", "count"))
    .reset_index()
)
unit_stats = unit_rej.merge(unit_total, on="委托单位", how="left")
unit_stats["退单率"] = unit_stats["退单笔数"] / unit_stats["总单量"]
# 按退单笔数由高到低排序（同笔数时按退单金额降序），取 Top20（上下两图共用该单位集合）
unit_stats = unit_stats.sort_values(
    ["退单笔数", "退单金额万元"], ascending=[False, False]
).reset_index(drop=True)
unit_stats_top = unit_stats.head(20)

if len(unit_stats_top) > 0:
    # 注：unit_stats_top 已按退单笔数降序排列，首位即退单体量最大单位
    top_volume_unit = unit_stats_top.iloc[0]["委托单位"]
    top_rate_unit = unit_stats_top.loc[unit_stats_top["退单率"].idxmax(), "委托单位"]
    top_amount_unit2 = unit_stats_top.loc[unit_stats_top["退单金额万元"].idxmax(), "委托单位"]
    st.markdown(
        f"**🚨 黑榜提示：【{top_volume_unit}】退单体量居首，"
        f"【{top_rate_unit}】退单率最高，"
        f"【{top_amount_unit2}】退单涉及金额最大**"
    )



    x_pos = np.arange(len(unit_stats_top))
    companies_list = unit_stats_top["委托单位"].tolist()
    counts_list = unit_stats_top["退单笔数"].tolist()
    amounts_list = unit_stats_top["退单金额万元"].tolist()

    # 上图：退单笔数 + 退单金额（双柱形图，双轴）
    fig3a, ax3a = plt.subplots(figsize=(12, 4.0))
    fig3a.patch.set_facecolor("#fafcff")
    ax3a.set_facecolor("#fafcff")

    bar_w = 0.38
    cnt_color = "#3A6EA1"
    amt_color = "#E67E22"

    bars_cnt = ax3a.bar(x_pos - bar_w / 2, counts_list, width=bar_w,
                        label="退单笔数（笔）", color=cnt_color,
                        alpha=0.85, edgecolor="white", linewidth=0.8, zorder=2)
    for bar, cnt in zip(bars_cnt, counts_list):
        ax3a.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.15,
                  str(cnt), ha="center", va="bottom", fontsize=7.5,
                  color=cnt_color, fontweight="bold")

    max_amt = max(amounts_list) if amounts_list else 1
    ax3a_r = ax3a.twinx()
    bars_amt = ax3a_r.bar(x_pos + bar_w / 2, amounts_list, width=bar_w,
                          label="退单涉及金额（万元）", color=amt_color,
                          alpha=0.85, edgecolor="white", linewidth=0.8, zorder=2)
    for bar, val in zip(bars_amt, amounts_list):
        if abs(val) > 0:
            ax3a_r.text(bar.get_x() + bar.get_width() / 2,
                        bar.get_height() + max_amt * 0.015,
                        f"{val:,.0f}", ha="center", va="bottom",
                        fontsize=7.5, color=amt_color, fontweight="bold")

    ax3a.set_ylabel("退单笔数（笔）", fontsize=9, color=cnt_color)
    ax3a.set_xticks(x_pos)
    ax3a.set_xticklabels(companies_list, rotation=25, ha="right", fontsize=7.5)
    ax3a.spines["top"].set_visible(False)
    ax3a.spines["right"].set_visible(False)
    ax3a.grid(axis="y", linestyle="--", alpha=0.35, zorder=0)
    ax3a.yaxis.set_major_locator(mticker.MaxNLocator(integer=True))
    ax3a.tick_params(axis="y", colors=cnt_color)

    ax3a_r.set_ylabel("退单涉及金额（万元）", fontsize=9, color=amt_color)
    ax3a_r.tick_params(axis="y", colors=amt_color)
    ax3a_r.spines["top"].set_visible(False)
    ax3a_r.spines["left"].set_visible(False)

    legend_handles3 = [
        MplRectangle((0, 0), 1, 1, facecolor=cnt_color, edgecolor="white", label="退单笔数（笔）"),
        MplRectangle((0, 0), 1, 1, facecolor=amt_color, edgecolor="white", label="退单涉及金额（万元）"),
    ]
    ax3a.legend(handles=legend_handles3, loc="upper right", fontsize=8, framealpha=0.7, edgecolor="#ccc")
    plt.tight_layout()
    st.pyplot(fig3a)
    plt.close(fig3a)

    # 下图：退单率（折线图 · 同一批 Top20，按退单率由高到低重新排序）
    unit_stats_rate_sorted = unit_stats_top.sort_values("退单率", ascending=False).reset_index(drop=True)
    rate_companies = unit_stats_rate_sorted["委托单位"].tolist()
    rates_sorted = unit_stats_rate_sorted["退单率"].tolist()
    n_rate = len(rate_companies)
    x_rate = np.arange(n_rate)

    fig3b, ax3b = plt.subplots(figsize=(12, 3.5))
    fig3b.patch.set_facecolor("#fafcff")
    ax3b.set_facecolor("#fafcff")

    rate_color = "#8E44AD"
    ax3b.plot(x_rate, rates_sorted, color=rate_color, linewidth=2.0,
              marker="o", markersize=5, markerfacecolor="white",
              markeredgewidth=1.8, markeredgecolor=rate_color, zorder=3)
    max_rate = max(rates_sorted) if rates_sorted else 1
    for i, rate in enumerate(rates_sorted):
        ax3b.text(i, rate + max_rate * 0.015, f"{rate:.1%}",
                  ha="center", va="bottom", fontsize=7.5, color=rate_color)

    ax3b.set_ylabel("退单率", fontsize=9, color=rate_color)
    ax3b.set_xticks(x_rate)
    ax3b.set_xticklabels(rate_companies, rotation=25, ha="right", fontsize=7.5)
    ax3b.yaxis.set_major_formatter(mticker.FuncFormatter(lambda y, _: f"{y:.0%}"))
    ax3b.spines["top"].set_visible(False)
    ax3b.spines["right"].set_visible(False)
    ax3b.grid(axis="y", linestyle="--", alpha=0.35, zorder=0)
    ax3b.tick_params(axis="y", colors=rate_color)
    ax3b.axhline(y=0.15, color="#E74C3C", linestyle="--", linewidth=1.5, alpha=0.85, zorder=4)
    ax3b.text(n_rate - 0.5, 0.155, "管理红线 (15%)",
              color="#E74C3C", fontsize=7.5, ha="right", va="bottom")
    plt.tight_layout()
    st.pyplot(fig3b)
    plt.close(fig3b)

else:
    st.info("暂无退单数据。")

# ============================================================
# 【Top 3 重点监管单位体检报告】四维加权风险评级
# ============================================================

st.markdown("---")
st.subheader("🏥 Top 3 重点监管单位体检报告（四维加权风险评级）")
st.info(
    "**📐 综合风险分计算逻辑：**\n\n"
    "**综合风险分 = 退单率得分 × 20% + 退单笔数得分 × 20% + 退单金额得分 × 40% + 标签集中率得分 × 20%**\n\n"
    "- **退单率得分**：该单位退单率 ÷ 所有单位最高退单率 × 100（满分100，越高表示退单率越严重）\n"
    "- **退单笔数得分**：该单位退单笔数 ÷ 所有单位最高退单笔数 × 100（满分100，越高表示退单体量越大、重复审核人耗越高）\n"
    "- **退单金额得分**：该单位退单金额 ÷ 所有单位最高退单金额 × 100（满分100，越高表示涉及金额越大）\n"
    "- **标签集中率得分**：该单位最高频退单标签占比 × 100（满分100，越高表示退单原因越集中、越易整改）\n\n"
    "**评级标准：** 综合风险分 ≥ 70 → 🔴 重点关注；40–69 → 🟡 一般关注；< 40 → 🟢 运行平稳"
)


if rejected_unique_count > 0:
    unit_rej2 = (
        df[df["is_rejected"]]
        .drop_duplicates(subset="单据编号", keep="first")
        .groupby("委托单位")
        .agg(退单笔数=("单据编号", "count"), 退单金额万元=("金额（万元）", "sum"))
        .reset_index()
    )
    unit_total2 = (
        df.drop_duplicates(subset="单据编号", keep="last")
        .groupby("委托单位")
        .agg(总单量=("单据编号", "count"))
        .reset_index()
    )
    unit_cc = unit_rej2.merge(unit_total2, on="委托单位", how="left")
    unit_cc["退单率"] = unit_cc["退单笔数"] / unit_cc["总单量"]

    def calc_label_concentration(unit_name):
        unit_rej_df = df_rej_records[df_rej_records["委托单位"] == unit_name]["处理意见标签化"].dropna()
        if len(unit_rej_df) == 0:
            return 0.0
        vc = unit_rej_df.value_counts()
        return float(vc.iloc[0]) / len(unit_rej_df)

    unit_cc["标签集中率"] = unit_cc["委托单位"].apply(calc_label_concentration)

    max_rate = unit_cc["退单率"].max() if unit_cc["退单率"].max() > 0 else 1
    max_count = unit_cc["退单笔数"].max() if unit_cc["退单笔数"].max() > 0 else 1
    max_amount = unit_cc["退单金额万元"].max() if unit_cc["退单金额万元"].max() > 0 else 1

    unit_cc["退单率得分"] = unit_cc["退单率"] / max_rate * 100
    unit_cc["退单笔数得分"] = unit_cc["退单笔数"] / max_count * 100
    unit_cc["金额得分"] = unit_cc["退单金额万元"].clip(lower=0) / max_amount * 100
    unit_cc["标签集中率得分"] = unit_cc["标签集中率"] * 100
    unit_cc["综合风险分"] = (
        unit_cc["退单率得分"] * 0.20 +
        unit_cc["退单笔数得分"] * 0.20 +
        unit_cc["金额得分"] * 0.40 +
        unit_cc["标签集中率得分"] * 0.20
    )


    unit_cc = unit_cc.sort_values(["综合风险分", "退单笔数"], ascending=[False, False]).reset_index(drop=True)
    top3_units = unit_cc.head(3)["委托单位"].tolist()

    DONUT_COLORS = ["#C0392B", "#2980B9", "#7FB3D5", "#E5E7E9"]
    top3_top1_reasons = []
    col_cards = st.columns(3)

    for idx, unit_name in enumerate(top3_units):
        with col_cards[idx]:
            unit_row = unit_cc[unit_cc["委托单位"] == unit_name].iloc[0]
            unit_rate = unit_row["退单率"]
            unit_count = int(unit_row["退单笔数"])
            unit_amount = unit_row["退单金额万元"]
            unit_conc = unit_row["标签集中率"]
            risk_score = unit_row["综合风险分"]


            if risk_score >= 70:
                rating_label = "🔴 重点关注"
                rating_color = "#c0392b"
            elif risk_score >= 40:
                rating_label = "🟡 一般关注"
                rating_color = "#e67e22"
            else:
                rating_label = "🟢 运行平稳"
                rating_color = "#27ae60"

            st.markdown(
                f"**{unit_name}**  \n"
                f"<span style='color:{rating_color}; font-weight:bold;'>{rating_label}</span>  \n"
                f"<small style='color:#888;'>"
                f"综合风险分：{risk_score:.1f} | 退单笔数：{unit_count}笔 | "
                f"退单率：{unit_rate:.1%} | 退单金额：{unit_amount:,.1f}万元 | "
                f"标签集中率：{unit_conc:.1%}"
                f"</small>",

                unsafe_allow_html=True,
            )

            unit_rej_df2 = df_rej_records[df_rej_records["委托单位"] == unit_name]
            tag_counts = unit_rej_df2["处理意见标签化"].dropna().value_counts()

            if len(tag_counts) > 0:
                if len(tag_counts) > 3:
                    top3_tags = tag_counts.iloc[:3]
                    others_sum = tag_counts.iloc[3:].sum()
                    labels_donut = top3_tags.index.tolist() + ["其他"]
                    sizes_donut = top3_tags.values.tolist() + [int(others_sum)]
                else:
                    labels_donut = tag_counts.index.tolist()
                    sizes_donut = tag_counts.values.tolist()

                n_slices = len(labels_donut)
                slice_colors = DONUT_COLORS[:n_slices]
                explode_d = [0.06 if i == 0 else 0 for i in range(n_slices)]

                fig_d, ax_d = plt.subplots(figsize=(3, 3))
                fig_d.patch.set_facecolor("none")
                ax_d.set_facecolor("none")

                pie_result = ax_d.pie(
                    sizes_donut, labels=None, colors=slice_colors, explode=explode_d,
                    autopct=lambda pct: f"{pct:.1f}%" if pct >= 6 else "",
                    pctdistance=0.75, startangle=90,
                    wedgeprops={"linewidth": 1.2, "edgecolor": "white"},
                )
                autotexts_d = pie_result[2] if type(pie_result) is tuple else getattr(pie_result, 'autotexts', [])
                for at in autotexts_d:
                    at.set_fontsize(8)
                    at.set_color("white")
                    at.set_fontweight("bold")

                centre_circle = MplCircle((0, 0), 0.52, fc="white", linewidth=0)
                ax_d.add_patch(centre_circle)
                ax_d.axis("equal")

                legend_labels_d = [lbl[:10] + "..." if len(lbl) > 10 else lbl for lbl in labels_donut]
                ax_d.legend(pie_result[0], legend_labels_d, loc="upper center",
                            bbox_to_anchor=(0.5, -0.05), ncol=1, frameon=False, fontsize=9, handlelength=1.5)
                plt.tight_layout(pad=0.3)
                st.pyplot(fig_d)
                plt.close(fig_d)
                top3_top1_reasons.append((unit_name, labels_donut[0]))
            else:
                st.info("该单位在当前筛选条件下暂无退单记录。")
                top3_top1_reasons.append((unit_name, "未知痛点"))

    # 督办令
    if len(top3_top1_reasons) >= 1:
        _names = [item[0] for item in top3_top1_reasons]
        _reasons = [item[1] for item in top3_top1_reasons]
        while len(_names) < 3:
            _names.append("—")
        while len(_reasons) < 3:
            _reasons.append("—")

        top_company_1, top_company_2, top_company_3 = _names[0], _names[1], _names[2]
        top_reason_1, top_reason_2, top_reason_3 = _reasons[0], _reasons[1], _reasons[2]

        if top_reason_2 == top_reason_3:
            if top_reason_1 == top_reason_2:
                order_text = (
                    f"经系统穿透，【{top_company_1}】、【{top_company_2}】与【{top_company_3}】"
                    f"的核心合规卡点高度一致，均集中在'{top_reason_1}'。"
                    f"建议应付组将上述痛点列入本月严控指标，"
                    f"针对性提供沟通服务。"
                )
            else:
                order_text = (
                    f"经系统穿透，【{top_company_1}】的核心合规卡点为'{top_reason_1}'，"
                    f"【{top_company_2}】与【{top_company_3}】的高频违规点均集中在'{top_reason_2}'。"
                    f"建议应付组将上述痛点列入本月严控指标，"
                    f"针对性提供沟通服务。"
                )
        elif top_reason_1 == top_reason_2:
            order_text = (
                f"经系统穿透，【{top_company_1}】与【{top_company_2}】的核心合规卡点均为'{top_reason_1}'，"
                f"【{top_company_3}】的高频违规点集中在'{top_reason_3}'。"
                f"建议应付组将上述痛点列入本月严控指标，"
                f"针对性提供沟通服务。"
            )
        elif top_reason_1 == top_reason_3:
            order_text = (
                f"经系统穿透，【{top_company_1}】与【{top_company_3}】的核心合规卡点均为'{top_reason_1}'，"
                f"【{top_company_2}】的高频违规点集中在'{top_reason_2}'。"
                f"建议应付组将上述痛点列入本月严控指标，"
                f"针对性提供沟通服务。"
            )
        else:
            order_text = (
                f"经系统穿透，【{top_company_1}】的核心合规卡点为'{top_reason_1}'，"
                f"【{top_company_2}】与【{top_company_3}】的高频违规点集中在"
                f"'{top_reason_2}'与'{top_reason_3}'。"
                f"建议应付组将上述痛点列入本月严控指标，"
                f"针对性提供沟通服务。"
            )

        st.error(f"🚨 **应付组合规攻坚重点**\n\n{order_text}")

else:
    st.info("暂无退单数据，请调整侧边栏筛选条件后重试。")

st.markdown("<br>", unsafe_allow_html=True)

# ============================================================
# 【新增图表模块】多维度专项分析
# ============================================================

st.markdown("---")
st.markdown("### 🔬 多维度专项分析")

new_col1, new_col2 = st.columns(2, gap="large")

with new_col1:
    st.markdown("#### 👤 处理人工作量 vs 退单量")
    handler_total = df.groupby("处理人")["单据编号"].count().reset_index()
    handler_total.columns = ["处理人", "总处理量"]
    handler_rej = df_rej_records.groupby("处理人")["单据编号"].count().reset_index()
    handler_rej.columns = ["处理人", "退单量"]
    handler_stats = handler_total.merge(handler_rej, on="处理人", how="left").fillna(0)
    handler_stats["退单量"] = handler_stats["退单量"].astype(int)
    handler_stats = handler_stats.sort_values("总处理量", ascending=False)

    x_h = np.arange(len(handler_stats))
    w = 0.38
    fig_h, ax_h = plt.subplots(figsize=(6.5, 4.0))
    fig_h.patch.set_facecolor("#fafcff")
    ax_h.set_facecolor("#fafcff")

    bars_h1 = ax_h.bar(x_h - w/2, handler_stats["总处理量"], width=w,
                        label="总处理量", color="#4A6FA5", alpha=0.85, edgecolor="white")
    bars_h2 = ax_h.bar(x_h + w/2, handler_stats["退单量"], width=w,
                        label="退单量", color="#C0392B", alpha=0.85, edgecolor="white")

    for bar in bars_h1:
        ax_h.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 5,
                  str(int(bar.get_height())), ha="center", va="bottom", fontsize=8)
    for bar in bars_h2:
        if bar.get_height() > 0:
            ax_h.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 5,
                      str(int(bar.get_height())), ha="center", va="bottom", fontsize=8, color="#C0392B")

    ax_h.set_xticks(x_h)
    ax_h.set_xticklabels(handler_stats["处理人"].tolist(), fontsize=9)
    ax_h.set_ylabel("笔数", fontsize=10)
    ax_h.legend(fontsize=9, framealpha=0.7)
    ax_h.spines["top"].set_visible(False)
    ax_h.spines["right"].set_visible(False)
    ax_h.grid(axis="y", linestyle="--", alpha=0.4)
    ax_h.yaxis.set_major_locator(mticker.MaxNLocator(integer=True))
    plt.tight_layout()
    st.pyplot(fig_h)
    plt.close(fig_h)

with new_col2:
    st.markdown("#### 🔄 交易类型退单分布")
    trade_rej = df_rej_records["交易类型"].value_counts().head(10)
    if len(trade_rej) > 0:
        t_labels = trade_rej.index.tolist()[::-1]
        t_values = trade_rej.values.tolist()[::-1]
        t_total = float(sum(t_values))
        cmap_t = plt.get_cmap("Oranges")
        t_colors = cmap_t(np.linspace(0.35, 0.85, len(t_labels)))

        fig_t, ax_t = plt.subplots(figsize=(6.5, 4.0))
        fig_t.patch.set_facecolor("#fafcff")
        ax_t.set_facecolor("#fafcff")

        bars_t = ax_t.barh(t_labels, t_values, color=t_colors, edgecolor="white", linewidth=0.8, height=0.6)
        for bar, val in zip(bars_t, t_values):
            pct = val / t_total * 100 if t_total > 0 else 0
            ax_t.text(bar.get_width() + 0.3, bar.get_y() + bar.get_height()/2,
                      f"{val} ({pct:.1f}%)", va="center", ha="left", fontsize=8, color="#333")

        ax_t.set_xlabel("退单笔数", fontsize=10)
        ax_t.spines["top"].set_visible(False)
        ax_t.spines["right"].set_visible(False)
        ax_t.grid(axis="x", linestyle="--", alpha=0.4)
        ax_t.xaxis.set_major_locator(mticker.MaxNLocator(integer=True))
        plt.tight_layout()
        st.pyplot(fig_t)
        plt.close(fig_t)
    else:
        st.info("暂无退单数据。")

new_col3, new_col4 = st.columns(2, gap="large")

with new_col3:
    st.markdown("#### 🔁 重复退单分析（重复退单次数）")
    rej_times = df[df["is_rejected"]].drop_duplicates(subset="单据编号", keep="first")["驳出共享次数"]
    rej_times_int = pd.to_numeric(
        rej_times.apply(lambda x: str(x).strip()).replace("--", np.nan),
        errors="coerce"
    ).dropna().astype(int)
    rej_dist = rej_times_int.value_counts().sort_index()

    if len(rej_dist) > 0:
        fig_r, ax_r = plt.subplots(figsize=(5.5, 4.0))
        fig_r.patch.set_facecolor("#fafcff")
        ax_r.set_facecolor("#fafcff")

        r_labels = [f"退回{k}次" for k in rej_dist.index]
        r_values = rej_dist.values.tolist()
        r_total = float(sum(r_values))
        DONUT_COLORS_R = ["#2980B9", "#E67E22", "#C0392B", "#8E44AD"][:len(r_labels)]
        explode_r = [0.06 if i == 0 else 0 for i in range(len(r_labels))]

        pie_r = ax_r.pie(
            r_values, labels=None, colors=DONUT_COLORS_R, explode=explode_r,
            autopct=lambda pct: f"{pct:.1f}%" if pct >= 5 else "",
            pctdistance=0.75, startangle=90,
            wedgeprops={"linewidth": 1.2, "edgecolor": "white"},
        )
        autotexts_r = pie_r[2] if type(pie_r) is tuple else getattr(pie_r, 'autotexts', [])
        for at in autotexts_r:
            at.set_fontsize(9)
            at.set_color("white")
            at.set_fontweight("bold")

        centre_r = MplCircle((0, 0), 0.52, fc="white", linewidth=0)
        ax_r.add_patch(centre_r)
        ax_r.axis("equal")

        legend_r = [f"{lbl}（{val}笔，{val/r_total*100:.1f}%）" for lbl, val in zip(r_labels, r_values)]
        ax_r.legend(pie_r[0], legend_r, loc="upper center", bbox_to_anchor=(0.5, -0.05),
                    ncol=1, frameon=False, fontsize=9)
        plt.tight_layout(pad=0.3)
        st.pyplot(fig_r)
        plt.close(fig_r)
    else:
        st.info("暂无重复退单数据。")

with new_col4:
    st.markdown("#### 🖥️ 来源系统退单对比")
    # 将"财务会计"映射为"应付管理"（仅用于展示，不修改原始df）
    df_sys_display = df.copy()
    df_sys_display["来源系统"] = df_sys_display["来源系统"].replace({"财务会计": "应付管理"})
    sys_total = df_sys_display.drop_duplicates(subset="单据编号", keep="last").groupby("来源系统")["单据编号"].count()
    sys_rej = df_sys_display[df_sys_display["is_rejected"]].drop_duplicates(subset="单据编号", keep="first").groupby("来源系统")["单据编号"].count()
    sys_stats = pd.DataFrame({"总单量": sys_total, "退单量": sys_rej}).fillna(0).astype(int)
    sys_stats["退单率"] = sys_stats["退单量"] / sys_stats["总单量"]
    sys_stats = sys_stats.reset_index()

    if len(sys_stats) > 0:
        x_s = np.arange(len(sys_stats))
        w_s = 0.35
        fig_s, ax_s = plt.subplots(figsize=(5.5, 4.0))
        fig_s.patch.set_facecolor("#fafcff")
        ax_s.set_facecolor("#fafcff")

        bars_s1 = ax_s.bar(x_s - w_s/2, sys_stats["总单量"], width=w_s,
                            label="总单量", color="#4A6FA5", alpha=0.85, edgecolor="white")
        bars_s2 = ax_s.bar(x_s + w_s/2, sys_stats["退单量"], width=w_s,
                            label="退单量", color="#C0392B", alpha=0.85, edgecolor="white")

        for bar in bars_s1:
            ax_s.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 10,
                      str(int(bar.get_height())), ha="center", va="bottom", fontsize=9)
        for bar in bars_s2:
            if bar.get_height() > 0:
                ax_s.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 10,
                          str(int(bar.get_height())), ha="center", va="bottom", fontsize=9, color="#C0392B")

        ax_s.set_xticks(x_s)
        ax_s.set_xticklabels(sys_stats["来源系统"].tolist(), fontsize=10)
        ax_s.set_ylabel("笔数", fontsize=10)
        ax_s.legend(fontsize=9, framealpha=0.7)
        ax_s.spines["top"].set_visible(False)
        ax_s.spines["right"].set_visible(False)
        ax_s.grid(axis="y", linestyle="--", alpha=0.4)
        ax_s.yaxis.set_major_locator(mticker.MaxNLocator(integer=True))

        ax_s_r = ax_s.twinx()
        ax_s_r.plot(x_s, sys_stats["退单率"], color="#8E44AD", linewidth=2,
                    marker="D", markersize=7, markerfacecolor="white",
                    markeredgewidth=2, markeredgecolor="#8E44AD", zorder=3)
        ax_s_r.yaxis.set_major_formatter(mticker.FuncFormatter(lambda y, _: f"{y:.1%}"))
        ax_s_r.set_ylabel("退单率", fontsize=10, color="#8E44AD")
        ax_s_r.tick_params(axis="y", colors="#8E44AD")
        ax_s_r.spines["top"].set_visible(False)
        ax_s_r.spines["left"].set_visible(False)
        plt.tight_layout()
        st.pyplot(fig_s)
        plt.close(fig_s)
    else:
        st.info("暂无数据。")

# 制单人退单Top15
st.markdown("#### 📝 制单人退单 Top 15 排名")
maker_rej = (
    df[df["is_rejected"]]
    .drop_duplicates(subset="单据编号", keep="first")
    .groupby(["制单人编码", "制单人"])["单据编号"]
    .count()
    .reset_index()
)
maker_rej.columns = ["制单人编码", "制单人", "退单笔数"]
maker_rej = maker_rej.sort_values("退单笔数", ascending=False).head(15)

if len(maker_rej) > 0:
    m_labels = [f"{row['制单人']}({row['制单人编码']})" for _, row in maker_rej.iterrows()]
    m_labels_rev = m_labels[::-1]
    m_values_rev = maker_rej["退单笔数"].tolist()[::-1]
    m_total = float(sum(m_values_rev))

    cmap_m = plt.get_cmap("Reds")
    m_colors = cmap_m(np.linspace(0.3, 0.8, len(m_labels_rev)))

    fig_m, ax_m = plt.subplots(figsize=(12, 5.0))
    fig_m.patch.set_facecolor("#fafcff")
    ax_m.set_facecolor("#fafcff")

    bars_m = ax_m.barh(m_labels_rev, m_values_rev, color=m_colors, edgecolor="white", linewidth=0.8, height=0.65)
    for bar, val in zip(bars_m, m_values_rev):
        pct = val / m_total * 100 if m_total > 0 else 0
        ax_m.text(bar.get_width() + 0.1, bar.get_y() + bar.get_height()/2,
                  f"{val} ({pct:.1f}%)", va="center", ha="left", fontsize=8.5, color="#333")

    ax_m.set_xlabel("退单笔数", fontsize=10)
    ax_m.spines["top"].set_visible(False)
    ax_m.spines["right"].set_visible(False)
    ax_m.grid(axis="x", linestyle="--", alpha=0.4)
    ax_m.xaxis.set_major_locator(mticker.MaxNLocator(integer=True))
    if m_values_rev:
        bars_m[-1].set_edgecolor("#c0392b")
        bars_m[-1].set_linewidth(2)
    plt.tight_layout()
    st.pyplot(fig_m)
    plt.close(fig_m)

    # ── 退单 Top5 制单人追溯到委托单位 ──
    st.markdown("##### 📋 退单 Top5 制单人追溯到委托单位")
    top5_makers = maker_rej.head(5)
    top5_maker_codes = top5_makers["制单人编码"].tolist()
    # 从退单记录中筛选Top5制单人，取每人出现最多的委托单位（众数），构建5行×3列表格
    df_rej_dedup = df[df["is_rejected"]].drop_duplicates(subset="单据编号", keep="first")
    rows_top5 = []
    for rank_i, (_, maker_row) in enumerate(top5_makers.iterrows(), start=1):
        code = maker_row["制单人编码"]
        name = maker_row["制单人"]
        sub = df_rej_dedup[df_rej_dedup["制单人编码"] == code]["委托单位"].dropna()
        top_unit = sub.value_counts().index[0] if len(sub) > 0 else "—"
        rows_top5.append({
            "排名": rank_i,
            "制单人（编码）": f"{name}（{code}）",
            "委托单位": top_unit,
        })
    df_top5_unit = pd.DataFrame(rows_top5)
    st.dataframe(
        df_top5_unit,
        use_container_width=True,
        height=215,
        column_config={
            "排名": st.column_config.NumberColumn(label="🏅 排名", width="small"),
            "制单人（编码）": st.column_config.TextColumn(label="👤 制单人（编码）", width="medium"),
            "委托单位": st.column_config.TextColumn(label="🏢 委托单位", width="large"),
        },
        hide_index=True,
    )
else:
    st.info("暂无制单人退单数据。")

# ============================================================
# 【BIP系统优化与攻坚进度追踪】模块（真实数据驱动）
# ============================================================

@st.cache_data
def load_bip_issues(bip_bytes):
    """读取BIP系统优化台账（分类明细sheet）"""
    df = pd.read_excel(bip_bytes, sheet_name="分类明细")
    return df

st.markdown("---")
st.subheader("🛠️ 用友 BIP 系统优化与攻坚追踪台账")

# ── BIP台账数据加载 ──
if bip_file is None:
    st.info(
        "💡 **BIP系统优化台账暂未加载**：请在左侧侧边栏上传 `BIP应付组系统优化点_分类汇总.xlsx`，"
        "即可查看系统优化攻坚进度追踪看板。"
    )
else:
    df_bip = load_bip_issues(bip_file)

    # ── 状态值标准化 ──
    STATUS_ORDER = ["已完成", "进行中", "待评估", "关闭需求"]
    STATUS_COLORS = {
        "已完成":   "#27ae60",
        "进行中":   "#2980b9",
        "待评估":   "#f39c12",
        "关闭需求": "#95a5a6",
    }
    # 未推进状态（攻坚盲区）
    PENDING_STATUSES = ["待评估"]

    # ── KPI 动态计算 ──
    total_bip = len(df_bip)
    closed_bip = int((df_bip["状态"] == "已完成").sum())
    closed_bip_pct = round(closed_bip / total_bip * 100) if total_bip > 0 else 0
    inprogress_bip = int((df_bip["状态"] == "进行中").sum())

    # ── KPI 卡片（3个）──
    bip_kpi1, bip_kpi2, bip_kpi3 = st.columns(3)
    with bip_kpi1:
        st.markdown(
            f"<div style='background:#f0f6ff; border-left:5px solid #2e6da4; padding:16px 20px; border-radius:8px; min-height:110px;'>"
            f"<div style='font-size:1.8rem;'>📌</div>"
            f"<div style='color:#555; font-size:0.82rem;'>累计提出系统需求</div>"
            f"<div style='color:#2e6da4; font-size:1.7rem; font-weight:bold;'>{total_bip} 项</div>"
            f"<div style='color:#888; font-size:0.78rem;'>覆盖全流程核心痛点</div></div>",
            unsafe_allow_html=True,
        )
    with bip_kpi2:
        st.markdown(
            f"<div style='background:#f0fff4; border-left:5px solid #27ae60; padding:16px 20px; border-radius:8px; min-height:110px;'>"
            f"<div style='font-size:1.8rem;'>✅</div>"
            f"<div style='color:#555; font-size:0.82rem;'>已闭环解决/完成</div>"
            f"<div style='color:#27ae60; font-size:1.7rem; font-weight:bold;'>{closed_bip} 项</div>"
            f"<div style='color:#888; font-size:0.78rem;'>整体进度约 {closed_bip_pct}%</div></div>",
            unsafe_allow_html=True,
        )
    with bip_kpi3:
        st.markdown(
            f"<div style='background:#fdf0ff; border-left:5px solid #8e44ad; padding:16px 20px; border-radius:8px; min-height:110px;'>"
            f"<div style='font-size:1.8rem;'>⚙️</div>"
            f"<div style='color:#555; font-size:0.82rem;'>当前攻坚/跟进中</div>"
            f"<div style='color:#8e44ad; font-size:1.7rem; font-weight:bold;'>{inprogress_bip} 项</div>"
            f"<div style='color:#888; font-size:0.78rem;'>持续推进，紧盯落地</div></div>",
            unsafe_allow_html=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)

    # ============================================================
    # 【攻坚态势图】
    # ============================================================
    st.markdown("#### 📊 系统优化攻坚态势图")

    # ── 数据预处理 ──
    all_statuses_bip = df_bip["状态"].dropna().unique().tolist()
    # 确保状态顺序一致
    ordered_statuses = [s for s in STATUS_ORDER if s in all_statuses_bip] + \
                       [s for s in all_statuses_bip if s not in STATUS_ORDER]

    # 一级分类 × 状态 交叉计数
    l1_cross = df_bip.groupby(["一级分类", "状态"]).size().unstack(fill_value=0)
    # 补全缺失状态列
    for s in ordered_statuses:
        if s not in l1_cross.columns:
            l1_cross[s] = 0
    l1_cross = l1_cross[ordered_statuses]
    # 按"未推进数量"降序排列（攻坚盲区最多的排最上）
    l1_cross["_pending"] = l1_cross[[s for s in PENDING_STATUSES if s in l1_cross.columns]].sum(axis=1)
    l1_cross = l1_cross.sort_values("_pending", ascending=True)  # ascending=True → 最多的在图表最上方
    l1_cross = l1_cross.drop(columns=["_pending"])

    # 解决方案分布
    solution_counts = df_bip["解决方案"].fillna("暂无方案").value_counts()

    # 二级分类未解决密度
    df_bip["_is_pending"] = df_bip["状态"].isin(PENDING_STATUSES)
    l2_pending = df_bip.groupby("二级分类").agg(
        未解决=("_is_pending", "sum"),
        总计=("_is_pending", "count")
    ).reset_index()
    l2_pending = l2_pending[l2_pending["未解决"] > 0].sort_values("未解决", ascending=True)

    # ── 图表布局：左右两列 ──
    bip_chart_col1, bip_chart_col2 = st.columns([3, 2], gap="large")

    with bip_chart_col1:
        st.markdown("##### 一级分类 · 攻坚进度（按状态堆叠）")
        categories = l1_cross.index.tolist()
        n_cats = len(categories)

        fig_bip1, ax_bip1 = plt.subplots(figsize=(8, max(3.5, n_cats * 0.75)))
        fig_bip1.patch.set_facecolor("#fafcff")
        ax_bip1.set_facecolor("#fafcff")

        lefts = np.zeros(n_cats)
        bar_handles = []
        for status in ordered_statuses:
            vals = l1_cross[status].values.astype(float)
            color = STATUS_COLORS.get(status, "#aaaaaa")
            bars_bip = ax_bip1.barh(
                categories, vals, left=lefts,
                color=color, edgecolor="white", linewidth=0.8,
                height=0.55, label=status,
            )
            # 在条形内标注数量（>0才显示）
            for i, (val, left) in enumerate(zip(vals, lefts)):
                if val > 0:
                    ax_bip1.text(
                        left + val / 2, i, str(int(val)),
                        ha="center", va="center", fontsize=8.5,
                        color="white", fontweight="bold",
                    )
            lefts += vals
            bar_handles.append(bars_bip[0])

        # 右侧标注完成率
        for i, cat in enumerate(categories):
            total_cat = l1_cross.loc[cat].sum()
            done_cat = l1_cross.loc[cat].get("已完成", 0)
            pct = done_cat / total_cat * 100 if total_cat > 0 else 0
            ax_bip1.text(
                lefts[i] + 0.3, i,
                f"完成率 {pct:.0f}%",
                va="center", ha="left", fontsize=8, color="#555",
            )

        ax_bip1.set_xlabel("需求数量（项）", fontsize=9)
        ax_bip1.xaxis.set_major_locator(mticker.MaxNLocator(integer=True))
        ax_bip1.spines["top"].set_visible(False)
        ax_bip1.spines["right"].set_visible(False)
        ax_bip1.grid(axis="x", linestyle="--", alpha=0.3)
        ax_bip1.legend(
            handles=bar_handles, labels=ordered_statuses,
            loc="lower right", fontsize=8, framealpha=0.7,
            ncol=min(3, len(ordered_statuses)),
        )
        plt.tight_layout()
        st.pyplot(fig_bip1)
        plt.close(fig_bip1)

    with bip_chart_col2:
        st.markdown("##### 解决路径 · 资源投入分布")
        sol_labels = solution_counts.index.tolist()
        sol_values = solution_counts.values.tolist()
        sol_total = float(sum(sol_values))

        SOLUTION_COLORS = {
            "配置":   "#27ae60",
            "客开":   "#e67e22",
            "集团管理员及共享各业务小组更新": "#e74c3c",
            "——":    "#95a5a6",
        }
        sol_colors = [SOLUTION_COLORS.get(lbl, "#aaaaaa") for lbl in sol_labels]
        explode_sol = [0.05 if i == 0 else 0 for i in range(len(sol_labels))]

        fig_bip2, ax_bip2 = plt.subplots(figsize=(4.5, 4.5))
        fig_bip2.patch.set_facecolor("#fafcff")
        ax_bip2.set_facecolor("#fafcff")

        pie_sol = ax_bip2.pie(
            sol_values, labels=None, colors=sol_colors, explode=explode_sol,
            autopct=lambda pct: f"{pct:.1f}%" if pct >= 5 else "",
            pctdistance=0.72, startangle=90,
            wedgeprops={"linewidth": 1.2, "edgecolor": "white"},
        )
        autotexts_sol = pie_sol[2] if type(pie_sol) is tuple else getattr(pie_sol, 'autotexts', [])
        for at in autotexts_sol:
            at.set_fontsize(9)
            at.set_color("white")
            at.set_fontweight("bold")

        centre_sol = MplCircle((0, 0), 0.50, fc="white", linewidth=0)
        ax_bip2.add_patch(centre_sol)
        ax_bip2.text(0, 0, f"共\n{total_bip}项", ha="center", va="center",
                     fontsize=10, fontweight="bold", color="#333")
        ax_bip2.axis("equal")

        legend_sol = [f"{lbl}（{val}项，{val/sol_total*100:.1f}%）"
                      for lbl, val in zip(sol_labels, sol_values)]
        ax_bip2.legend(
            pie_sol[0], legend_sol,
            loc="upper center", bbox_to_anchor=(0.5, -0.04),
            ncol=1, frameon=False, fontsize=8.5, handlelength=1.5,
        )
        plt.tight_layout(pad=0.5)
        st.pyplot(fig_bip2)
        plt.close(fig_bip2)

    # ── 二级分类未解决密度排行（全宽）──
    if len(l2_pending) > 0:
        st.markdown("##### 二级分类 · 攻坚盲区密度排行")
        st.caption('仅统计"待评估"的需求，代表当前尚未推进完成的攻坚盲区')

        n_l2 = len(l2_pending)
        cmap_l2 = plt.get_cmap("Reds")
        l2_colors = cmap_l2(np.linspace(0.35, 0.80, n_l2))

        fig_bip3, ax_bip3 = plt.subplots(figsize=(10, max(3.0, n_l2 * 0.55)))
        fig_bip3.patch.set_facecolor("#fafcff")
        ax_bip3.set_facecolor("#fafcff")

        bars_l2 = ax_bip3.barh(
            l2_pending["二级分类"], l2_pending["未解决"],
            color=l2_colors, edgecolor="white", linewidth=0.8, height=0.55,
        )
        for bar, row in zip(bars_l2, l2_pending.itertuples()):
            ax_bip3.text(
                bar.get_width() + 0.1,
                bar.get_y() + bar.get_height() / 2,
                f"{int(row.未解决)} 项未推进 / 共 {int(row.总计)} 项",
                va="center", ha="left", fontsize=8.5, color="#555",
            )

        ax_bip3.set_xlabel("未推进需求数量（项）", fontsize=9)
        ax_bip3.xaxis.set_major_locator(mticker.MaxNLocator(integer=True))
        ax_bip3.spines["top"].set_visible(False)
        ax_bip3.spines["right"].set_visible(False)
        ax_bip3.grid(axis="x", linestyle="--", alpha=0.3)
        plt.tight_layout()
        st.pyplot(fig_bip3)
        plt.close(fig_bip3)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── 状态筛选器 + 明细表 ──
    all_bip_statuses = df_bip["状态"].dropna().unique().tolist()
    selected_bip_statuses = st.multiselect(
        "🔎 按进度状态筛选（可多选）",
        options=all_bip_statuses, default=all_bip_statuses,
        help="选择一个或多个进度状态，过滤下方明细表",
        key="bip_status_filter",
    )
    if not selected_bip_statuses:
        selected_bip_statuses = all_bip_statuses

    df_bip_filtered = df_bip[df_bip["状态"].isin(selected_bip_statuses)].copy()
    # 移除辅助列
    if "_is_pending" in df_bip_filtered.columns:
        df_bip_filtered = df_bip_filtered.drop(columns=["_is_pending"])

    # 状态颜色高亮函数
    def highlight_bip_status(val):
        color_map = {
            "已完成":   "background-color: #e8f5e9; color: #1a7a3a;",
            "进行中":   "background-color: #e3f0fb; color: #1a5a8a;",
            "待评估":   "background-color: #fff8e1; color: #8a6000;",
            "暂无状态": "background-color: #f5f5f5; color: #666;",
            "关闭需求": "background-color: #fce8e8; color: #8a2020;",
        }
        return color_map.get(val, "")

    st.dataframe(
        df_bip_filtered.style.map(highlight_bip_status, subset=["状态"]),
        use_container_width=True,
        height=520,
        column_config={
            "序号": st.column_config.NumberColumn(label="序号", width="small"),
            "BIP系统模块": st.column_config.TextColumn(label="🖥️ BIP系统模块", width="medium"),
            "BIP系统优化点": st.column_config.TextColumn(label="📝 BIP系统优化点", width="large"),
            "状态": st.column_config.TextColumn(label="📊 状态", width="small"),
            "一级分类": st.column_config.TextColumn(label="🏷️ 一级分类", width="small"),
            "二级分类": st.column_config.TextColumn(label="🔖 二级分类", width="small"),
            "解决方案": st.column_config.TextColumn(label="💡 解决方案", width="medium"),
        },
        hide_index=True,
    )
    st.markdown(
        f"<small style='color:#888;'>共展示 {len(df_bip_filtered)} 条 / 台账合计 {len(df_bip)} 条系统优化需求</small>",
        unsafe_allow_html=True,
    )

st.markdown("<br>", unsafe_allow_html=True)

# ============================================================
# 【底部数据表】全量作业明细
# ============================================================

st.markdown("### 📋 作业明细底层数据（全字段预览）")

def highlight_status(val):
    if val == "审核驳回":
        return "background-color: #fde8e8"
    elif val == "审核通过":
        return "background-color: #e8f5e9"
    return ""

exclude_cols = {"处理时间_dt", "月份", "年份", "作业时长_秒", "is_rejected", "金额（元）"}
display_cols_all = [c for c in df.columns if c not in exclude_cols]
df_display = df[display_cols_all].copy()

show_rows = st.slider(
    "📌 展示行数",
    min_value=10, max_value=min(200, len(df_display)),
    value=min(50, len(df_display)), step=10,
    help="拖动滑块控制底部数据表展示的行数",
)

st.dataframe(
    df_display.head(show_rows).style.map(highlight_status, subset=["作业状态"]),
    use_container_width=True,
    height=420,
)

st.markdown(
    f"<small style='color:#888;'>共展示 {show_rows} 条 / 筛选后合计 {len(df)} 条作业记录（含通过与驳回）</small>",
    unsafe_allow_html=True,
)

# ============================================================
# 【页脚】
# ============================================================
st.markdown("---")
st.markdown(
    "<div style='text-align:center; color:#aaa; font-size:0.8rem;'>"
    "财务共享中心应付组 · 退单分析自动化看板 &nbsp;|&nbsp; "
    "数据来源：财务共享中心BIP系统导出（2026年1月-6月）&nbsp;|&nbsp; Powered by Streamlit"
    "</div>",
    unsafe_allow_html=True,
)

# -*- coding: utf-8 -*-
"""
generate_sample_data.py — 生成脱敏样例数据（字段与真实 BIP 导出数据完全对齐）

用途：
  1) 供看板 demo / 面试演示使用（无需真实敏感数据）
  2) 供公开分享（GitHub）使用——所有单位、人员均为虚构示例，无真实敏感信息

输出（samples/ 目录）：
  - 样例_主表_作业明细查询列表.xlsx / .csv   （主表 45 列）
  - 样例_金额表_作业明细查询-金额.xlsx / .csv （金额表 25 列）
  - 样例_BIP系统优化点_分类汇总.xlsx / .csv   （BIP 台账 7 列，读取【分类明细】sheet）

运行方式：
  python tools/generate_sample_data.py
"""
import os
import random
import numpy as np
import pandas as pd
from datetime import datetime, timedelta

SEED = 42
random.seed(SEED)
np.random.seed(SEED)

BASE = os.path.dirname(os.path.abspath(__file__))
SAMPLES = os.path.normpath(os.path.join(BASE, "..", "samples"))
os.makedirs(SAMPLES, exist_ok=True)

# ============================================================
# 业务参数（全部为虚构脱敏数据）
# ============================================================
N_MAIN = 2000         # 主表作业明细条数
N_BIP = 30            # BIP 系统优化需求条数

# 委托单位 20 家（帕累托分布：前 5 家合计约 53%，尾部单位体量小、退单率易偏高）
UNITS = [
    ("A01", "示例集团·甲公司"), ("A02", "示例集团·乙公司"),
    ("A03", "示例集团·丙公司"), ("A04", "示例集团·丁公司"),
    ("A05", "示例集团·戊公司"), ("A06", "示例集团·己公司"),
    ("A07", "示例集团·庚公司"), ("A08", "示例集团·辛公司"),
    ("A09", "示例集团·壬公司"), ("A10", "示例集团·癸公司"),
    ("A11", "示例集团·子公司甲"), ("A12", "示例集团·子公司乙"),
    ("A13", "示例集团·子公司丙"), ("A14", "示例集团·子公司丁"),
    ("A15", "示例集团·子公司戊"), ("A16", "示例集团·子公司己"),
    ("A17", "示例集团·子公司庚"), ("A18", "示例集团·子公司辛"),
    ("A19", "示例集团·子公司壬"), ("A20", "示例集团·子公司癸"),
]
UNIT_WEIGHTS = [0.14, 0.12, 0.10, 0.09, 0.08, 0.07, 0.06, 0.05, 0.045, 0.04,
                0.035, 0.03, 0.027, 0.024, 0.022, 0.02, 0.017, 0.014, 0.011, 0.009]

HANDLERS = [("AUD01", "审核员01"), ("AUD02", "审核员02"), ("AUD03", "审核员03"),
            ("AUD04", "审核员04"), ("AUD05", "审核员05"), ("AUD06", "审核员06"),
            ("AUD07", "审核员07")]

MAKERS = [("MAK01", "制单人01"), ("MAK02", "制单人02"), ("MAK03", "制单人03"),
          ("MAK04", "制单人04"), ("MAK05", "制单人05"), ("MAK06", "制单人06"),
          ("MAK07", "制单人07"), ("MAK08", "制单人08"), ("MAK09", "制单人09"),
          ("MAK10", "制单人10"), ("MAK11", "制单人11"), ("MAK12", "制单人12"),
          ("MAK13", "制单人13"), ("MAK14", "制单人14"), ("MAK15", "制单人15"),
          ("MAK16", "制单人16")]

TRADE_TYPES = ["应付付款单（材料采购）", "应付付款单（服务费）", "应付付款单（差旅费）",
               "通用报销单（费用报销）", "应付确认单（物料类）", "其他专项付款单"]
DOC_TYPES = ["付款单", "通用报销单", "应付发票"]

LABELS = ["附件不完整", "附件信息不一致", "未关联合同", "业务类型填写错误", "特殊原因", "其他"]
LABEL_WEIGHTS = [0.38, 0.20, 0.16, 0.12, 0.08, 0.06]

SOURCE_SYS = ["财务会计", "费控服务"]
SYS_WEIGHTS = [0.72, 0.28]

# 月度业务量分布（1-6 月：1月/6月冲量、中段略低，让月度趋势图有涨落）
MONTH_WEIGHTS = [0.20, 0.18, 0.16, 0.15, 0.13, 0.18]

def rand_dt():
    """随机时间戳，范围为 2026-01-01 ~ 2026-06-30（按月度业务量权重分布）"""
    month = int(np.random.choice([1, 2, 3, 4, 5, 6], p=MONTH_WEIGHTS))
    day = random.randint(1, 28)
    h = random.randint(8, 20)
    m = random.randint(0, 59)
    s = random.randint(0, 59)
    return datetime(2026, month, day).replace(hour=h, minute=m, second=s)


def rand_duration():
    return f"{random.randint(1, 40)}分{random.randint(0, 59)}秒"


def rand_seconds():
    return f"{random.randint(10, 300)}秒"


# ============================================================
# 生成主表（45 列，与真实字段对齐）
# ============================================================
def build_main():
    rows = []
    shared_center = "示例集团财务共享中心（本部）"

    for i in range(1, N_MAIN + 1):
        doc_no = f"AP-2026-{i:06d}"
        src_sys = np.random.choice(SOURCE_SYS, p=SYS_WEIGHTS)
        unit_code, unit_name = random.choices(UNITS, weights=UNIT_WEIGHTS, k=1)[0]
        maker_code, maker_name = random.choices(MAKERS, k=1)[0]
        handler_code, handler_name = random.choices(HANDLERS, k=1)[0]
        trade_type = np.random.choice(TRADE_TYPES)
        doc_type = np.random.choice(DOC_TYPES)

        doc_date = rand_dt()
        create_dt = rand_dt()
        pool_dt = create_dt + timedelta(seconds=random.randint(1, 60))
        in_group_dt = pool_dt + timedelta(seconds=random.randint(1, 120))
        handle_dt = in_group_dt + timedelta(minutes=random.randint(1, 200))
        # 处理时间不越过 6 月 30 日
        if handle_dt > datetime(2026, 6, 30, 23, 59, 59):
            handle_dt = datetime(2026, 6, 30, 23, 0, 0) + timedelta(seconds=random.randint(0, 3500))

        # 退单逻辑：约 10% 最终驳回；另有约 5.5% 最终通过但曾被退回 → 口径②约 15%
        is_final_reject = random.random() < 0.10
        rejected_before = is_final_reject or (random.random() < 0.055)
        times = 0 if not rejected_before else (random.randint(1, 3) if is_final_reject else 1)

        if rejected_before:
            label = np.random.choice(LABELS, p=LABEL_WEIGHTS)
        else:
            label = np.nan

        if is_final_reject:
            status = "审核驳回"
            opinion = "驳回"
            reject_type, reject_mode = "驳回", 1
            reject_link = np.random.choice(["制单人", "前端单位"])
            reject_recv = maker_name if reject_link == "制单人" else unit_name
        else:
            status = "审核通过"
            opinion = "同意"
            reject_type, reject_mode, reject_link, reject_recv = np.nan, 0, np.nan, np.nan

        rows.append({
            "来源系统": src_sys,
            "单据编号": doc_no,
            "单据日期": doc_date.strftime("%Y-%m-%d"),
            "单据类型": doc_type,
            "交易类型": trade_type,
            "委托单位编码": unit_code,
            "委托单位": unit_name,
            "制单人编码": maker_code,
            "制单人": maker_name,
            "附件个数": np.nan,
            "影像张数": random.randint(0, 8),
            "共享中心": shared_center,
            "共享服务类型": "财务共享",
            "服务目录": trade_type,
            "任务属性": "通用审批类",
            "创建时间": create_dt.strftime("%Y-%m-%d %H:%M:%S"),
            "入池时间": pool_dt.strftime("%Y-%m-%d %H:%M:%S"),
            "入组时间": in_group_dt.strftime("%Y-%m-%d %H:%M:%S"),
            "任务类型": "FSSC审核岗",
            "任务量系数": 1,
            "优先级": np.nan,
            "工作组": "应付成本组",
            "派单到人模式": "自动派单",
            "派单到人规则": "已派任务少（优先派）",
            "提取/派单时间": in_group_dt.strftime("%Y-%m-%d %H:%M:%S"),
            "操作终端": np.random.choice(["PC端", "移动端"], p=[0.55, 0.45]),
            "处理人编码": handler_code,
            "处理人": handler_name,
            "处理时间": handle_dt.strftime("%Y-%m-%d %H:%M:%S"),
            "处理意见": opinion,
            "处理意见标签化": label,
            "作业状态": status,
            "作业时长": rand_duration(),
            "修正作业时长": rand_duration() if (rejected_before and random.random() < 0.8) else "--",
            "超期标记": 0,
            "超期时间": np.nan,
            "共享是否处理完成": "是",
            "共享处理次数": 1 if times == 0 else times,
            "进入共享次数": 1,
            "页面审核时长": rand_seconds(),
            "驳回类型": reject_type,
            "驳回模式": reject_mode,
            "驳回接收环节": reject_link,
            "驳回接收人": reject_recv,
            "驳出共享次数": "--" if times == 0 else times,
        })

    cols = ["来源系统", "单据编号", "单据日期", "单据类型", "交易类型", "委托单位编码", "委托单位",
            "制单人编码", "制单人", "附件个数", "影像张数", "共享中心", "共享服务类型", "服务目录",
            "任务属性", "创建时间", "入池时间", "入组时间", "任务类型", "任务量系数", "优先级",
            "工作组", "派单到人模式", "派单到人规则", "提取/派单时间", "操作终端", "处理人编码",
            "处理人", "处理时间", "处理意见", "处理意见标签化", "作业状态", "作业时长", "修正作业时长",
            "超期标记", "超期时间", "共享是否处理完成", "共享处理次数", "进入共享次数", "页面审核时长",
            "驳回类型", "驳回模式", "驳回接收环节", "驳回接收人", "驳出共享次数"]
    return pd.DataFrame(rows, columns=cols)


# ============================================================
# 生成金额表（25 列，与真实字段对齐）
# ============================================================
def build_amount(main_df: pd.DataFrame):
    rows = []
    shared_center = "示例集团财务共享中心（本部）"
    for _, r in main_df.iterrows():
        # 约 15% 单据无金额记录（对应真实业务中部分业务无金额）
        if random.random() < 0.15:
            continue
        n_lines = random.randint(1, 3)
        pool_dt = pd.to_datetime(r["入池时间"])
        for _ in range(n_lines):
            # 对数正态分布：多数单据小额（几千~几十万）、少数大额（可达数百万）
            amt = float(np.clip(np.random.lognormal(mean=11.0, sigma=1.4), 500, 5_000_000))
            rows.append({
                "委托单位编码": r["委托单位编码"],
                "委托单位名称": r["委托单位"],
                "单据编号": r["单据编号"],
                "单据类型": r["单据类型"],
                "交易类型": r["交易类型"],
                "币种": "人民币",
                "金额": round(amt, 2),
                "优先级": np.nan,
                "超期状态": "--",
                "单据日期": r["单据日期"],
                "任务类型": "FSSC审核岗",
                "工作组": "应付成本组",
                "处理人": r["处理人"],
                "作业状态": r["作业状态"],
                "入池时间": r["入池时间"],
                "入组时间": r["入组时间"],
                "提取时间": (pool_dt + timedelta(seconds=random.randint(10, 600))).strftime("%Y-%m-%d %H:%M:%S"),
                "通过时间": "--" if r["作业状态"] == "审核驳回" else
                             (pd.to_datetime(r["处理时间"]) + timedelta(minutes=random.randint(1, 120))).strftime("%Y-%m-%d %H:%M:%S"),
                "是否全电发票": np.nan,
                "收单状态": "--",
                "实物状态": np.nan,
                "共享中心": shared_center,
                "服务目录": r["交易类型"],
                "自动审批状态": "非自动审批",
                "任务量": 1,
            })

    cols = ["委托单位编码", "委托单位名称", "单据编号", "单据类型", "交易类型", "币种", "金额",
            "优先级", "超期状态", "单据日期", "任务类型", "工作组", "处理人", "作业状态",
            "入池时间", "入组时间", "提取时间", "通过时间", "是否全电发票", "收单状态", "实物状态",
            "共享中心", "服务目录", "自动审批状态", "任务量"]
    return pd.DataFrame(rows, columns=cols)


# ============================================================
# 生成 BIP 系统优化台账（7 列，与真实字段对齐）
# ============================================================
def build_bip():
    """生成 BIP 系统优化台账——使用独立随机流，不受主表样本量影响，内容恒定可复现"""
    bip_np = np.random.RandomState(SEED)   # 独立 NumPy RNG（与主表/金额表随机流隔离）

    modules = ["应付组", "费控模块", "发票模块", "合同模块"]
    l1_cats = ["管控类", "操作类", "系统类"]
    l2_cats = ["预算管控", "合同管控", "流程优化", "字段校验", "效率提升", "数据治理"]
    solutions = ["配置", "客开", "集团管理员及共享各业务小组更新", "暂无方案", "——"]
    statuses = ["已完成", "进行中", "待评估", "关闭需求", "暂无状态"]
    status_weights = [0.33, 0.20, 0.27, 0.13, 0.07]
    topics = ["审批环节", "影像管理", "预算控制", "合同关联", "单据校验", "派单规则"]

    rows = []
    for i in range(1, N_BIP + 1):
        rows.append({
            "序号": i,
            "BIP系统模块": bip_np.choice(modules),
            "BIP系统优化点": f"示例优化需求{i:02d}：{bip_np.choice(topics)}相关的流程与系统功能改进建议",
            "状态": bip_np.choice(statuses, p=status_weights),
            "一级分类": bip_np.choice(l1_cats, p=[0.40, 0.35, 0.25]),
            "二级分类": bip_np.choice(l2_cats),
            "解决方案": bip_np.choice(solutions, p=[0.30, 0.15, 0.20, 0.25, 0.10]),
        })
    return pd.DataFrame(rows, columns=["序号", "BIP系统模块", "BIP系统优化点", "状态", "一级分类", "二级分类", "解决方案"])


def save(df: pd.DataFrame, name: str, sheet_name: str = "Sheet1"):
    xlsx = os.path.join(SAMPLES, f"样例_{name}.xlsx")
    csv = os.path.join(SAMPLES, f"样例_{name}.csv")
    df.to_excel(xlsx, index=False, engine="openpyxl", sheet_name=sheet_name)
    df.to_csv(csv, index=False, encoding="utf-8-sig")
    return xlsx, csv


def main():
    print("[1/3] 生成主表 ...")
    main_df = build_main()
    rej_cnt = int(pd.to_numeric(main_df["驳出共享次数"].astype(str).replace("--", "0"), errors="coerce").fillna(0).ge(1).sum())
    print(f"      {len(main_df)} 条记录，曾被退回(口径②) {rej_cnt} 条，退单率 {rej_cnt/len(main_df):.1%}")

    print("[2/3] 生成金额表 ...")
    amount_df = build_amount(main_df)
    print(f"      {len(amount_df)} 条金额行")

    for df, name in [(main_df, "主表_作业明细查询列表"), (amount_df, "金额表_作业明细查询-金额")]:
        x, c = save(df, name, sheet_name="Sheet1")
        print(f"  OK -> {x}")
        print(f"  OK -> {c}")

    # BIP 台账：若已存在则跳过重写（保持既有内容不变）；首次生成时写入【分类明细】sheet
    bip_x = os.path.join(SAMPLES, "样例_BIP系统优化点_分类汇总.xlsx")
    bip_c = os.path.join(SAMPLES, "样例_BIP系统优化点_分类汇总.csv")
    if os.path.exists(bip_x) and os.path.exists(bip_c):
        print("[3/3] BIP 台账已存在，跳过重写（内容保持不变）")
    else:
        print("[3/3] 生成 BIP 台账 ...")
        bip_df = build_bip()
        print(f"      {len(bip_df)} 条优化需求")
        x, c = save(bip_df, "BIP系统优化点_分类汇总", sheet_name="分类明细")
        print(f"  OK -> {x}")
        print(f"  OK -> {c}")

    print("\n全部样例数据已生成，目录：", SAMPLES)


if __name__ == "__main__":
    main()

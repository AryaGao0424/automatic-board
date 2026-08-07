# -*- coding: utf-8 -*-
"""纯位置调整：将【Top 3 重点监管单位体检报告】板块移动至【各委托单位退单分析】之后、【多维度专项分析】之前。不修改任何代码字符。"""
import io
import sys

p = r"d:\@ZUEL\Intership\01-CJCY\01- Pay Group\automatic board\fssc_dashboard.py"

with io.open(p, "r", encoding="utf-8") as f:
    ls = f.readlines()

total_before = len(ls)
print("total lines before:", total_before)

# ---- 边界断言（基于用户提供的行号）----
# 待移动块：line890 ~ line1091 (0-based index 889..1091)
blk = ls[889:1092]
assert total_before == 1434, "unexpected total lines: %d" % total_before
assert "Top 3" in ls[890], "block header mismatch"
assert "Top 3" in ls[894], "subheader mismatch"
assert "应付组合规攻坚重点" in ls[1086], "order text missing"
assert ls[1090].strip().startswith("st.markdown"), "block end mismatch"
assert "# ===========" in ls[1092], "tail block header mismatch (multidim)"
assert "# ===========" in ls[633], "multidim header reference mismatch"

# 目标插入点：line631 空行之后、line632 多维度注释头之前 (0-based index 631)
assert ls[630].strip().startswith("st.info"), "line630 should be 各委托单位退单分析 end"
assert ls[631].strip() == "", "line631 should be blank"
assert "# ===========" in ls[632], "line632 should be 多维度注释头"

# ---- 拼接：去掉待移动块，然后在 index 631 处插入 ----
remaining = ls[:889] + ls[1092:]
out = remaining[:631] + blk + remaining[631:]

total_after = len(out)
print("total lines after:", total_after)
assert total_after == total_before, "line count changed!"

with io.open(p, "w", encoding="utf-8") as f:
    f.writelines(out)

print("MOVE OK")

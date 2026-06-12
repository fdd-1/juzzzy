#!/usr/bin/env python3
"""海外LP群发违规检查 - 相似度过滤 + 续费场景统计.

用法:
    python check.py <违规明细.xlsx> --baseline <报备文本.xlsx> [--output <输出.xlsx>] [--threshold 0.90]
"""
import argparse
import re
import sys
from datetime import date
from difflib import SequenceMatcher
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

try:
    import pandas as pd
except ImportError:
    print("缺少依赖，请先安装: pip install pandas openpyxl")
    sys.exit(1)

DEFAULT_THRESHOLD = 0.90
KEEP_SCENE = "续费"
RENEW_RATIO_WARN = 0.05  # 续费占比 < 5% 视为异常

# 明细表中可能出现的列名（按优先级匹配）
GROUP_COL_CANDIDATES = ["小组", "组别", "团队"]
LP_COL_CANDIDATES = ["LP姓名", "LP", "姓名", "班主任"]
SCENE_COL_CANDIDATES = ["场景", "违规场景", "群发场景"]
TEXT_COL_CANDIDATES = ["违规文本", "消息内容", "群发内容", "文本", "内容"]


def pick_col(df, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    # 模糊匹配
    for col in df.columns:
        for c in candidates:
            if c in str(col):
                return col
    return None


def read_excel_smart(path):
    """尝试多个 header 行读取，返回最合理的 DataFrame."""
    for header in [0, 1, 2, 3]:
        try:
            df = pd.read_excel(path, header=header)
            if pick_col(df, LP_COL_CANDIDATES) and pick_col(df, TEXT_COL_CANDIDATES):
                return df, header
        except Exception:
            continue
    # 兜底
    return pd.read_excel(path, header=0), 0


def load_baseline_texts(baseline_path):
    """从报备文本表中提取所有非空字符串作为候选池."""
    df = pd.read_excel(baseline_path, header=0, sheet_name=None)  # 读所有 sheet
    texts = set()
    for sheet_name, sheet_df in df.items():
        for col in sheet_df.columns:
            for val in sheet_df[col].dropna():
                s = str(val).strip()
                if len(s) >= 8:  # 太短的不作为模板
                    texts.add(s)
    return list(texts)


def normalize_text(s):
    """归一化文本：去空白、标点，便于相似度比较."""
    if not isinstance(s, str):
        s = str(s)
    s = re.sub(r"\s+", "", s)
    s = re.sub(r"[，。！？、,.!?:;：；\"'\"\"''（）()【】\[\]\-——~`]+", "", s)
    return s.lower()


def max_similarity(text, baseline_texts):
    """返回 text 与 baseline_texts 中最相似一条的相似度."""
    target = normalize_text(text)
    if not target:
        return 0.0
    best = 0.0
    for b in baseline_texts:
        b_norm = normalize_text(b)
        if not b_norm:
            continue
        # 长度差距过大直接跳过
        if abs(len(target) - len(b_norm)) > max(len(target), len(b_norm)) * 0.6:
            continue
        ratio = SequenceMatcher(None, target, b_norm).ratio()
        if ratio > best:
            best = ratio
            if best >= 0.99:
                break
    return best


def _self_check(counts: dict, detail_rows: int) -> bool:
    """打印自检表 + 校验，返回是否通过."""
    print("\n[自检] 数据链路核对")
    order = [
        "BI 原始行",
        "有效违规",
        "续费场景",
        "唯一文本",
        "命中报备(剔除)",
        "未命中",
        "同LP同文本去重后",
        "最终统计 LP 数",
        "最终违规合计",
    ]
    for k in order:
        if k in counts:
            print(f"  {k:<20} {counts[k]}")

    valid = counts.get("有效违规", 0)
    renew = counts.get("续费场景", 0)
    final = counts.get("最终违规合计", 0)
    final_lp = counts.get("最终统计 LP 数", 0)

    # 1) 任意一步为 0 → 中止
    zero_keys = [k for k in ("BI 原始行", "有效违规", "续费场景") if counts.get(k, 0) == 0]
    if zero_keys:
        print(f"  ✗ 异常：{zero_keys} 为 0，链路不通，请排查 BI 报表 / 列名映射")
        return False

    # 2) 续费占比 < 5% → 异常（场景列可能没识别到「续费」）
    if valid > 0:
        ratio = renew / valid
        if ratio < RENEW_RATIO_WARN:
            print(f"  ✗ 异常：续费占比 {ratio:.1%} < {RENEW_RATIO_WARN:.0%}，"
                  f"场景列可能未正确识别「续费」")
            return False

    # 3) 最终违规合计 != 违规明细行数 → 中止
    if final != detail_rows:
        print(f"  ✗ 异常：最终违规合计 {final} != 违规明细行数 {detail_rows}，"
              f"统计与明细不一致")
        return False

    # 4) 最终违规合计 > 0 但 LP 数 == 0 → 中止
    if final > 0 and final_lp == 0:
        print(f"  ✗ 异常：违规 {final} 条但 LP 数为 0，groupby 出错")
        return False

    print("  ✓ 链路一致")
    return True


def check(detail_path, baseline_path, output_path, threshold=DEFAULT_THRESHOLD):
    counts = {}  # 自检计数器
    print(f"[1] 读取违规明细: {detail_path}")
    df, header = read_excel_smart(detail_path)
    counts["BI 原始行"] = len(df)
    print(f"    ✓ header 行={header}, 共 {len(df)} 行, 列: {list(df.columns)[:10]}...")

    group_col = pick_col(df, GROUP_COL_CANDIDATES)
    lp_col = pick_col(df, LP_COL_CANDIDATES)
    scene_col = pick_col(df, SCENE_COL_CANDIDATES)
    text_col = pick_col(df, TEXT_COL_CANDIDATES)

    missing = [n for n, c in [("小组", group_col), ("LP姓名", lp_col),
                              ("场景", scene_col), ("违规文本", text_col)] if not c]
    if missing:
        print(f"    ✗ 缺少列: {missing}")
        print(f"    实际列: {list(df.columns)}")
        sys.exit(1)
    print(f"    ✓ 列匹配: 小组={group_col}, LP={lp_col}, 场景={scene_col}, 文本={text_col}")

    # 过滤掉无 LP 行
    df = df[df[lp_col].notna() & df[text_col].notna()].copy()
    df = df[~df[lp_col].astype(str).str.contains(r"总计|小计|合计", regex=True, na=False)]
    counts["有效违规"] = len(df)
    counts["剔除空LP/合计行"] = counts["BI 原始行"] - counts["有效违规"]
    print(f"    ✓ 有效违规记录: {len(df)}")

    # 仅保留续费场景
    df_renew = df[df[scene_col].astype(str).str.contains(KEEP_SCENE, na=False)].copy()
    counts["续费场景"] = len(df_renew)
    counts["剔除非续费"] = len(df) - len(df_renew)
    print(f"    ✓ 续费场景: {len(df_renew)} 条（剔除转介绍等 {len(df) - len(df_renew)} 条）")

    if df_renew.empty:
        print("    ⚠ 无续费违规，退出")
        return None

    print(f"[2] 加载报备文本: {baseline_path}")
    baseline_texts = load_baseline_texts(baseline_path)
    counts["报备文本数"] = len(baseline_texts)
    print(f"    ✓ 报备文本候选 {len(baseline_texts)} 条")

    print(f"[3] 计算相似度（阈值 {threshold:.0%}）...")
    # 先对所有违规文本归一化并去重，每个唯一文本只算一次
    df_renew["_norm_text"] = df_renew[text_col].apply(lambda t: normalize_text(str(t)))
    unique_norms = df_renew["_norm_text"].dropna().unique().tolist()
    counts["唯一文本"] = len(unique_norms)
    print(f"    ✓ 唯一文本数: {len(unique_norms)}（原 {len(df_renew)} 条）")

    # 预归一化报备文本池
    baseline_norms = [normalize_text(b) for b in baseline_texts]
    baseline_norms = [b for b in baseline_norms if b]

    # 缓存每个唯一文本的最高相似度
    sim_cache = {}
    total_unique = len(unique_norms)
    for i, target in enumerate(unique_norms, 1):
        if i % 200 == 0 or i == total_unique:
            print(f"    进度 {i}/{total_unique}", flush=True)
        if not target:
            sim_cache[target] = 0.0
            continue
        best = 0.0
        tlen = len(target)
        for b in baseline_norms:
            blen = len(b)
            if abs(tlen - blen) > max(tlen, blen) * 0.6:
                continue
            ratio = SequenceMatcher(None, target, b).ratio()
            if ratio > best:
                best = ratio
                if best >= 0.99:
                    break
        sim_cache[target] = best

    df_renew["_sim"] = df_renew["_norm_text"].map(sim_cache).fillna(0.0)
    df_renew["_is_reported"] = df_renew["_sim"] >= threshold
    counts["命中报备(剔除)"] = int(df_renew["_is_reported"].sum())
    counts["未命中"] = int((~df_renew["_is_reported"]).sum())
    print(f"    ✓ 命中报备: {counts['命中报备(剔除)']} 条")
    print(f"    ✓ 仍属违规: {counts['未命中']} 条")

    # 剔除报备命中
    df_violation = df_renew[~df_renew["_is_reported"]].copy()

    # 相同文本去重：同一LP的相同违规文本只算1次（已用归一化后的文本）
    before_dedup = len(df_violation)
    df_violation = df_violation.drop_duplicates(subset=[lp_col, "_norm_text"])
    counts["同LP同文本去重后"] = len(df_violation)
    counts["去重剔除"] = before_dedup - len(df_violation)
    print(f"    ✓ 文本去重: {before_dedup} → {len(df_violation)} 条（相同LP的相同文本只算1次）")

    # 按 LP 统计
    print("[4] 按LP统计...")
    stat = (
        df_violation.groupby([group_col, lp_col])
        .size()
        .reset_index(name="续费违规条数")
        .rename(columns={group_col: "小组", lp_col: "LP姓名"})
        .sort_values(["小组", "续费违规条数"], ascending=[True, False])
    )
    counts["最终统计 LP 数"] = len(stat)
    counts["最终违规合计"] = int(stat["续费违规条数"].sum()) if not stat.empty else 0
    print(f"    ✓ 共 {len(stat)} 位LP有续费违规")

    # 自检
    if not _self_check(counts, len(df_violation)):
        print("    ✗ 自检未通过，终止流程，不写出文件、不播报")
        sys.exit(2)

    # 输出
    print(f"[5] 写入: {output_path}")
    with pd.ExcelWriter(output_path, engine="openpyxl") as w:
        stat.to_excel(w, sheet_name="违规统计", index=False)
        # 详细条目（含相似度，方便人工复核）
        detail_out = df_violation[[group_col, lp_col, scene_col, text_col, "_sim"]].rename(
            columns={group_col: "小组", lp_col: "LP姓名", scene_col: "场景",
                     text_col: "违规文本", "_sim": "最高相似度"}
        )
        detail_out.to_excel(w, sheet_name="违规明细", index=False)
        # 命中报备的也单独一个 sheet 留底
        reported_out = df_renew[df_renew["_is_reported"]][[group_col, lp_col, scene_col, text_col, "_sim"]].rename(
            columns={group_col: "小组", lp_col: "LP姓名", scene_col: "场景",
                     text_col: "命中文本", "_sim": "相似度"}
        )
        reported_out.to_excel(w, sheet_name="命中报备(已剔除)", index=False)
    print(f"    ✓ 已写入 {output_path}")
    return output_path


def main():
    p = argparse.ArgumentParser(description="海外LP群发违规检查")
    p.add_argument("detail", help="海外LP群发违规明细 Excel")
    p.add_argument("--baseline", required=True, help="报备文本 Excel（钉钉文档导出）")
    p.add_argument("--output", help="输出统计 Excel 路径")
    p.add_argument("--threshold", type=float, default=DEFAULT_THRESHOLD, help="相似度阈值 0~1")
    args = p.parse_args()

    detail = Path(args.detail)
    baseline = Path(args.baseline)
    if not detail.exists():
        print(f"错误: 明细不存在 {detail}")
        sys.exit(1)
    if not baseline.exists():
        print(f"错误: 报备文本不存在 {baseline}")
        print(f"     请先把钉钉文档导出为 Excel 放到 {baseline}")
        sys.exit(1)

    output = args.output or str(detail.parent / f"违规统计_{date.today().strftime('%Y%m%d')}.xlsx")
    check(str(detail), str(baseline), output, args.threshold)


if __name__ == "__main__":
    main()

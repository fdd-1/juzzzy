"""通用数据格式化模块

规则:
1. 列头含"占比"、"率"等字眼 → 转百分比格式（保留两位小数）
2. 列头最后字含"加和" → 不转百分比，保留两位小数
3. 列头最后字含"数" → 整数格式，不处理
4. 其他数值列 → 保留两位小数

特殊规则:
- 删除空白列（全部为NaN的列）
- 删除空白行（全部为NaN的行）
- 删除口径说明行
"""
from __future__ import annotations
import sys
import pandas as pd
import re


def is_percent_column(col_name: str) -> bool:
    """判断是否是百分比列(列头含'率'或'占比')但不含'加和'。"""
    s = str(col_name)
    if '加和' in s:
        return False
    return '率' in s or '占比' in s or '比例' in s


def is_count_column(col_name: str) -> bool:
    """判断是否是计数列(列头最后字是'数'或含'人数'/'学员数'等)。"""
    s = str(col_name).strip()
    if not s:
        return False
    # 最后一个字是"数"
    if s.endswith('数'):
        return True
    # 包含"次数"/"人数"等
    count_keywords = ['次数', '人数', '学员数', '新生数', '通时', '消息数', 'leads数', '订单数']
    return any(kw in s for kw in count_keywords)


def is_jiahuo_column(col_name: str) -> bool:
    """判断是否是加和列(列头含'加和')。"""
    return '加和' in str(col_name)


def format_value(val, col_name: str):
    """根据列名格式化单个值。

    Returns:
        格式化后的值（百分比字符串/小数字符串/整数）
    """
    if pd.isna(val) or val is None or val == "":
        return ""

    # 尝试转数字
    try:
        if isinstance(val, str):
            # 已经是百分比字符串，保留
            if val.endswith('%'):
                return val
            num = float(val.replace(',', ''))
        else:
            num = float(val)
    except (ValueError, TypeError):
        return val  # 非数值，原样返回

    # 列头是百分比列
    if is_percent_column(col_name) and not is_jiahuo_column(col_name):
        # 0-1范围内 → ×100，否则保持
        if 0 <= num <= 1:
            return f"{num * 100:.2f}%"
        else:
            return f"{num:.2f}%"

    # 加和列 → 保留两位小数（不带%）
    if is_jiahuo_column(col_name):
        return round(num, 2)

    # 计数列 → 整数
    if is_count_column(col_name):
        return int(num) if num == int(num) else round(num, 2)

    # 其他数值 → 保留两位小数
    return round(num, 2)


def format_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """对整个 DataFrame 应用格式化。

    Returns:
        格式化后的 DataFrame
    """
    df = df.copy()

    # 1. 删除全空列
    df = df.dropna(axis=1, how='all')

    # 2. 删除全空行
    df = df.dropna(how='all')

    # 3. 处理重复列名
    if df.columns.duplicated().any():
        # 给重复列名添加后缀
        cols = []
        seen = {}
        for c in df.columns:
            if c in seen:
                seen[c] += 1
                cols.append(f"{c}_{seen[c]}")
            else:
                seen[c] = 0
                cols.append(c)
        df.columns = cols

    # 4. 对每列应用格式化
    for col in df.columns:
        df[col] = df[col].apply(lambda v: format_value(v, col))

    return df


def remove_empty_and_caliber(df: pd.DataFrame, key_col: str = None) -> pd.DataFrame:
    """删除空白列、空白行、口径行。

    Args:
        df: DataFrame
        key_col: 关键列（用于判断口径行的位置）
    """
    df = df.copy()

    # 1. 删除全空列
    df = df.dropna(axis=1, how='all')

    # 2. 删除全空行
    df = df.dropna(how='all')

    # 3. 删除口径行
    if key_col and key_col in df.columns:
        # 找到第一个口径行的位置
        cut_idx = None
        for i, val in enumerate(df[key_col]):
            if pd.isna(val):
                continue
            s = str(val)
            if any(kw in s for kw in ['口径说明', '口径：', '注：', '备注', '说明：']):
                cut_idx = i
                break
            # 数字开头的注释行 (1)、2、 等)
            if re.match(r'^\d+[\)、）.]', s):
                cut_idx = i
                break

        if cut_idx is not None:
            df = df.iloc[:cut_idx].reset_index(drop=True)

    # 4. 再次删除可能产生的空白行（key_col 为空但其他列也都为空的）
    df = df.dropna(how='all').reset_index(drop=True)

    return df


if __name__ == "__main__":
    # 测试
    print("测试列名识别:")
    test_cols = [
        "停课占比",
        "执行率加和",
        "学员数",
        "外呼跟进率",
        "首课跟进率",
        "新生数",
        "总通时_min",
        "leads数",
    ]
    for c in test_cols:
        is_pct = is_percent_column(c)
        is_cnt = is_count_column(c)
        is_jh = is_jiahuo_column(c)
        print(f"  {c}: 百分比={is_pct}, 计数={is_cnt}, 加和={is_jh}")

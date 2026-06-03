#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
海外益智服务池学员自动化处理工具

功能：
1. 从 BI 下载「海外思维续费规划表_新版_26年启用」报表
2. 按条件筛选（是否可续学员=1，月初是否续费=空白）
3. 与学员流转文件匹配，计算最终归属LP和小组
4. 创建六一标签和用户群
5. 同步标签数据到豌豆数仓
"""

import os
import sys
import json
import argparse
import subprocess
import logging
from datetime import datetime, timedelta
from pathlib import Path
import pandas as pd
import glob

# 配置日志
def setup_logger(log_dir):
    """配置日志"""
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, f"service_pool_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")

    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )
    return logging.getLogger(__name__)

def load_config(config_path="config.json"):
    """加载配置文件"""
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"配置文件不存在: {config_path}")

    with open(config_path, 'r', encoding='utf-8') as f:
        return json.load(f)

def create_monthly_workdir(base_dir, year_month):
    """创建月度工作目录"""
    dirs = {
        "downloads": os.path.join(base_dir, "data", "downloads"),
        "processed": os.path.join(base_dir, "data", "processed"),
        "reference": os.path.join(base_dir, "data", "reference"),
        "logs": os.path.join(base_dir, "logs")
    }

    for d in dirs.values():
        os.makedirs(d, exist_ok=True)

    return dirs

def download_bi_report(config, output_dir, year_month, logger):
    """
    使用 smartbi_browser_export.py 下载 BI 报表

    Args:
        config: 配置字典
        output_dir: 输出目录
        year_month: 年月字符串 (YYYY-MM)
        logger: 日志记录器

    Returns:
        下载的文件路径
    """
    logger.info(f"[1/7] 开始下载 BI 报表...")

    # 计算日期
    date_obj = datetime.strptime(year_month, "%Y-%m")
    first_day_of_month = date_obj.replace(day=1).strftime("%Y-%m-%d")
    last_day_of_prev_month = (date_obj.replace(day=1) - timedelta(days=1)).strftime("%Y-%m-%d")

    logger.info(f"  开课M计算时间: {first_day_of_month}")
    logger.info(f"  退费结束时间: {last_day_of_prev_month}")

    # 使用 smartbi_browser_export.py
    script_path = config.get("smartbi_browser_export_path",
        r"C:\Users\fengjianyi\Desktop\smartbi-data-cli-internal-20260526\smartbi-data-cli-internal-20260526\scripts\smartbi_browser_export.py")

    report_id = config["bi_report"]["report_id"]
    output_filename = f"服务池学员_{year_month.replace('-', '')}.xlsx"
    output_path = os.path.join(output_dir, output_filename)

    # 构建筛选条件 JSON
    filters = [
        ["开课M计算时间", first_day_of_month, first_day_of_month],
        ["退费结束时间", last_day_of_prev_month, last_day_of_prev_month],
        ["池子节点3", "服务月", "服务月"]
    ]
    filters_json = json.dumps(filters, ensure_ascii=False)

    # 构建命令
    cmd = [
        "python", script_path,
        "--report-id", report_id,
        "--output", output_path,
        "--max-rows", "10000",
        "--filters-json", filters_json,
        "--username", "76218",
        "--password", "123456"
    ]

    logger.info(f"  执行浏览器导出（筛选条件：{len(filters)}个）")

    env = os.environ.copy()
    env['PYTHONIOENCODING'] = 'utf-8'

    result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='ignore', env=env)

    if result.returncode != 0:
        logger.error(f"BI 报表下载失败: {result.stderr}")
        raise Exception(f"BI 报表下载失败: {result.stderr}")

    logger.info(f"  BI 报表下载完成")

    if not os.path.exists(output_path):
        raise Exception(f"未找到下载的文件: {output_path}")

    logger.info(f"  找到下载文件: {output_path}")

    return output_path

def filter_bi_data(bi_file_path, config, logger):
    """
    对 BI 报表进行二次筛选

    Args:
        bi_file_path: BI 报表文件路径
        config: 配置字典
        logger: 日志记录器

    Returns:
        筛选后的 DataFrame
    """
    logger.info(f"[2/7] 开始数据筛选...")

    # SIMPLE_REPORT 格式：第7行是列名，从第8行开始是数据
    df = pd.read_excel(bi_file_path, header=6)
    logger.info(f"  原始数据行数: {len(df)}")
    logger.info(f"  列数: {len(df.columns)}")

    # 检查关键列是否存在
    required_cols = ["是否可续学员", "月初是否续费", "学员ID", "大账号ID", "LP姓名", "LP组别"]
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        logger.error(f"实际列名: {list(df.columns)[:20]}")
        raise Exception(f"缺少必要列: {missing_cols}")

    # 筛选条件
    filter_cond = config["filter_conditions"]
    condition = (df["是否可续学员"] == filter_cond["是否可续学员"]) & \
                (df["月初是否续费"].isna() | (df["月初是否续费"] == ""))

    filtered = df[condition].copy()
    logger.info(f"  筛选后数据行数: {len(filtered)}")
    logger.info(f"  筛选条件: 是否可续学员={filter_cond['是否可续学员']}, 月初是否续费=空白")

    return filtered

def calculate_final_attribution(bi_df, liuzhuang_file_path, logger):
    """
    计算最终归属LP和归属小组

    匹配规则：
    1. 使用 BI 报表的「大账号ID」与学员流转文件的「学员id」匹配
    2. 如果匹配到：
       - 最终归属LP = 学员流转文件的「新班主任姓名」
       - 最终归属小组 = 学员流转文件的「新班主任组别」
    3. 如果未匹配到：
       - 最终归属LP = BI 报表的「LP姓名」
       - 最终归属小组 = BI 报表的「LP组别」

    Args:
        bi_df: BI 报表 DataFrame
        liuzhuang_file_path: 学员流转文件路径
        logger: 日志记录器

    Returns:
        添加了最终归属列的 DataFrame
    """
    logger.info(f"[3/7] 开始数据匹配...")

    if not os.path.exists(liuzhuang_file_path):
        raise Exception(f"学员流转文件不存在: {liuzhuang_file_path}")

    # 读取学员流转文件
    liuzhuang_df = pd.read_excel(liuzhuang_file_path)
    logger.info(f"  学员流转文件行数: {len(liuzhuang_df)}")

    # 检查必要列
    required_cols = ["学员id", "新班主任姓名", "新班主任组别"]
    missing_cols = [col for col in required_cols if col not in liuzhuang_df.columns]
    if missing_cols:
        raise Exception(f"学员流转文件缺少必要列: {missing_cols}")

    # 构建匹配字典（学员id -> 新LP信息）
    liuzhuang_dict = {}
    for _, row in liuzhuang_df.iterrows():
        student_id = row["学员id"]
        if pd.notna(student_id):
            liuzhuang_dict[int(student_id)] = {
                "新班主任姓名": row["新班主任姓名"],
                "新班主任组别": row["新班主任组别"]
            }

    logger.info(f"  学员流转字典条目数: {len(liuzhuang_dict)}")

    # 计算最终归属（使用大账号ID匹配）
    matched_count = 0
    unmatched_count = 0

    final_lp_list = []
    final_group_list = []

    for _, row in bi_df.iterrows():
        dadou_id = row["大账号ID"]
        if pd.notna(dadou_id) and int(dadou_id) in liuzhuang_dict:
            # 匹配到
            lp_info = liuzhuang_dict[int(dadou_id)]
            final_lp_list.append(lp_info["新班主任姓名"])
            final_group_list.append(lp_info["新班主任组别"])
            matched_count += 1
        else:
            # 未匹配到，使用原LP
            final_lp_list.append(row["LP姓名"])
            final_group_list.append(row["LP组别"])
            unmatched_count += 1

    bi_df["最终归属LP"] = final_lp_list
    bi_df["最终归属小组"] = final_group_list

    logger.info(f"  匹配到学员流转: {matched_count} 条")
    logger.info(f"  未匹配（使用原LP）: {unmatched_count} 条")

    return bi_df

def save_final_result(df, output_path, logger):
    """
    保存最终结果

    输出文件：
    - {output_path}: 完整数据（Sheet1 完整数据 + Sheet2 大账号ID）
    - {output_path 同目录}/dadou_ids_{YYYYMM}.xlsx: 单独的大账号ID文件（用于六一标签上传）

    Args:
        df: 最终 DataFrame
        output_path: 输出文件路径
        logger: 日志记录器

    Returns:
        dadou_excel_path: 单独的大账号ID文件路径（用于六一标签）
    """
    logger.info(f"[4/7] 保存处理结果...")

    # 准备大账号ID数据
    dadou_df = df[["大账号ID"]].copy()
    dadou_df = dadou_df[dadou_df["大账号ID"].notna()]
    dadou_df["大账号ID"] = dadou_df["大账号ID"].astype(int)

    # 保存完整结果（双 Sheet）
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Sheet1', index=False)
        logger.info(f"  Sheet1: 完整数据 ({len(df)} 行, {len(df.columns)} 列)")

        dadou_df.to_excel(writer, sheet_name='六一标签数据', index=False)
        logger.info(f"  Sheet2: 大账号ID ({len(dadou_df)} 条)")

    logger.info(f"  完整文件已保存: {output_path}")

    # 单独保存大账号ID文件（用于六一标签上传 - create.py 默认读第一个 Sheet）
    output_dir = os.path.dirname(output_path)
    base_name = os.path.basename(output_path)
    # 从文件名提取月份后缀
    import re
    m = re.search(r'(\d{6})', base_name)
    suffix = m.group(1) if m else datetime.now().strftime("%Y%m")
    dadou_excel_path = os.path.join(output_dir, f"dadou_ids_{suffix}.xlsx")
    dadou_df.to_excel(dadou_excel_path, index=False)
    logger.info(f"  大账号ID独立文件: {dadou_excel_path}")

    return dadou_excel_path

def create_liuyi_tag_and_group(excel_path, month_str, config, logger, dry_run=False):
    """
    创建六一标签和用户群

    Args:
        excel_path: 包含大账户ID的 Excel 文件路径（第一列是大账号ID）
        month_str: 月份字符串（如 "6月"）
        config: 配置字典
        logger: 日志记录器
        dry_run: 仅预览不执行

    Returns:
        执行结果字典 {tag_name, tag_id, group_id, group_name}
    """
    logger.info(f"[5/7] 创建六一标签和用户群...")

    tag_config = config["tag_config"]
    tag_name = f"{tag_config['prefix']}{config['year']}{month_str}{tag_config['suffix']}"

    logger.info(f"  标签名称: {tag_name}")
    logger.info(f"  用户群名称: {tag_name}")

    if dry_run:
        logger.info("  [DRY RUN] 跳过六一工作台操作")
        return {"tag_name": tag_name, "dry_run": True}

    liuyi_create_path = config["liuyi_create_path"]

    cmd = [
        "python",
        liuyi_create_path,
        "--input", excel_path,
        "--id-type", "wandou",
        "--tag-name", tag_name,
        "--group-name", tag_name
    ]

    if tag_config.get("with_wechat", False):
        cmd.append("--with-wechat")

    logger.info(f"  执行命令: python create.py --input ... --tag-name {tag_name}")

    env = os.environ.copy()
    env['PYTHONIOENCODING'] = 'utf-8'
    result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='ignore', env=env)

    if result.returncode != 0:
        logger.error(f"六一标签创建失败:\n{result.stdout}\n{result.stderr}")
        raise Exception(f"六一标签创建失败")

    # 打印 create.py 的输出
    for line in result.stdout.splitlines():
        if line.strip():
            logger.info(f"  {line}")

    # 从 create.py 同目录读取最新的 result_*.json，获取 tag_id 和 group_id
    liuyi_dir = Path(liuyi_create_path).parent
    result_jsons = sorted(liuyi_dir.glob("result_*.json"), key=lambda p: p.stat().st_mtime, reverse=True)
    info = {"tag_name": tag_name}
    if result_jsons:
        try:
            with open(result_jsons[0], 'r', encoding='utf-8') as f:
                info.update(json.load(f))
            logger.info(f"  tag_id: {info.get('tag_id')}, group_id: {info.get('group_id')}")
        except Exception as e:
            logger.warning(f"  读取 result.json 失败: {e}")

    return info

def sync_tag_to_warehouse(group_name, config, logger, dry_run=False):
    """
    同步标签数据到豌豆数仓

    Args:
        group_name: 用户群名称
        config: 配置字典
        logger: 日志记录器
        dry_run: 仅预览不执行

    Returns:
        执行结果
    """
    logger.info(f"[6/7] 同步标签数据到豌豆数仓...")
    logger.info(f"  用户群名称: {group_name}")

    if dry_run:
        logger.info("  [DRY RUN] 跳过数仓同步操作")
        return {"group_name": group_name, "dry_run": True}

    # 构建命令
    liuyi_sync_path = config["liuyi_sync_path"]

    cmd = [
        "python",
        liuyi_sync_path,
        "--user-group-name", group_name
    ]

    logger.info(f"  执行命令: {' '.join(cmd)}")

    env = os.environ.copy()
    env['PYTHONIOENCODING'] = 'utf-8'
    result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='ignore', env=env)

    if result.returncode != 0:
        logger.error(f"数仓同步失败: {result.stderr}")
        raise Exception(f"数仓同步失败: {result.stderr}")

    logger.info(f"  数仓同步完成")
    logger.info(f"  输出: {result.stdout[:200]}...")

    return {"group_name": group_name, "stdout": result.stdout}

def main():
    """主流程"""
    # 解析命令行参数
    parser = argparse.ArgumentParser(description='海外益智服务池学员自动化处理工具')
    parser.add_argument('--month', type=str, default=None,
                        help='处理月份（格式：YYYY-MM，默认当前月）')
    parser.add_argument('--liuzhuang-file', type=str, default=None,
                        help='学员流转文件路径（可选，默认从配置读取）')
    parser.add_argument('--dry-run', action='store_true',
                        help='仅预览不执行（跳过六一工作台操作）')
    parser.add_argument('--skip-liuyi', action='store_true',
                        help='跳过六一工作台操作（仅生成数据）')
    parser.add_argument('--config', type=str, default='config.json',
                        help='配置文件路径（默认：config.json）')

    args = parser.parse_args()

    # 加载配置
    script_dir = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(script_dir, args.config) if not os.path.isabs(args.config) else args.config
    config = load_config(config_path)

    # 确定处理月份
    year_month = args.month if args.month else config.get("current_month", datetime.now().strftime("%Y-%m"))

    # 创建工作目录
    work_dir = create_monthly_workdir(config["work_dir_base"], year_month)

    # 配置日志
    logger = setup_logger(work_dir["logs"])
    logger.info("=" * 60)
    logger.info(f"海外益智服务池学员自动化处理 - {year_month}")
    logger.info("=" * 60)

    try:
        # 1. 下载 BI 报表
        bi_file = download_bi_report(config, work_dir["downloads"], year_month, logger)

        # 2. 二次筛选
        filtered_df = filter_bi_data(bi_file, config, logger)

        # 3. 数据匹配
        liuzhuang_file = args.liuzhuang_file if args.liuzhuang_file else config["liuzhuang_file"]
        final_df = calculate_final_attribution(filtered_df, liuzhuang_file, logger)

        # 4. 保存结果（同时生成大账号ID独立文件）
        output_filename = f"服务池学员_{year_month.replace('-', '')}.xlsx"
        output_path = os.path.join(work_dir["processed"], output_filename)
        dadou_excel_path = save_final_result(final_df, output_path, logger)

        # 5. 创建六一标签和用户群（使用大账号ID独立文件）
        if not args.skip_liuyi:
            month_str = datetime.strptime(year_month, "%Y-%m").strftime("%-m月" if os.name != 'nt' else "%#m月")
            create_liuyi_tag_and_group(dadou_excel_path, month_str, config, logger, dry_run=args.dry_run)

            # 6. 同步到豌豆数仓
            tag_config = config["tag_config"]
            group_name = f"{tag_config['prefix']}{config['year']}{month_str}{tag_config['suffix']}"
            sync_tag_to_warehouse(group_name, config, logger, dry_run=args.dry_run)
        else:
            logger.info("[5/7] 跳过六一标签和用户群创建（--skip-liuyi）")
            logger.info("[6/7] 跳过数仓同步（--skip-liuyi）")

        # 7. 完成
        logger.info("=" * 60)
        logger.info("[7/7] 处理完成！")
        logger.info(f"处理结果已保存至: {output_path}")
        logger.info(f"共处理学员: {len(final_df)} 人")

        # 统计信息
        matched_count = (final_df["最终归属LP"] != final_df["LP姓名"]).sum()
        logger.info(f"匹配到学员流转: {matched_count} 人")
        logger.info(f"使用原LP: {len(final_df) - matched_count} 人")
        logger.info("=" * 60)

        return 0

    except Exception as e:
        logger.error("=" * 60)
        logger.error(f"处理失败: {str(e)}")
        logger.error("=" * 60)
        import traceback
        logger.error(traceback.format_exc())
        return 1

if __name__ == "__main__":
    sys.exit(main())


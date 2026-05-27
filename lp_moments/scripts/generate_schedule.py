"""
LP企微朋友圈排期批量生成脚本
调用Claude API，根据地区画像和产品卖点生成朋友圈排期与话术
"""

import os
import json
import argparse
from datetime import datetime
from pathlib import Path

try:
    import anthropic
except ImportError:
    print("请先安装 anthropic SDK: pip install anthropic")
    exit(1)

BASE_DIR = Path(__file__).resolve().parent.parent
CONTENT_DIR = BASE_DIR / "content_library"
OUTPUT_DIR = BASE_DIR / "output"


def read_file(path: Path) -> str:
    with open(path, "r", encoding="utf-8") as f:
        return f.read()


def load_content_library():
    return {
        "region_profiles": read_file(CONTENT_DIR / "region_profiles.md"),
        "product_selling_points": read_file(CONTENT_DIR / "product_selling_points.md"),
        "upgrade_season_guide": read_file(CONTENT_DIR / "upgrade_season_guide.md"),
    }


def build_prompt(region, stage, period, focus, custom_event, content_lib):
    region_label = "港澳台" if region == "hk" else "欧美澳"
    period_label = "一周" if period == "week" else "一个月"
    focus_map = {
        "upgrade": "升阶季转化",
        "retention": "留存维护",
        "renewal": "续费促进",
        "activity": "活动推广",
    }
    focus_label = focus_map.get(focus, focus)

    stage_instruction = f"目标阶段: {stage}" if stage != "all" else "覆盖全阶段(S1-S9)"

    event_note = f"\n当前特殊节点: {custom_event}" if custom_event else ""

    prompt = f"""你是一位资深的K12教育行业企微运营专家。请为LP（学管）生成{period_label}的企微朋友圈排期和话术。

## 基本要求
- 地区: {region_label}
- {stage_instruction}
- 营销重点: {focus_label}
- 排期周期: {period_label}{event_note}

## 地区用户画像
{content_lib['region_profiles']}

## 产品卖点素材
{content_lib['product_selling_points']}

## 升阶季参考（如适用）
{content_lib['upgrade_season_guide']}

## 输出要求
按以下节奏规划朋友圈排期：
- 情感共鸣帖（1-2条）：贴近家长焦虑/期望，建立信任
- 产品价值帖（1-2条）：突出课程卖点，对应地区诉求
- 学员成果帖（1条）：真实案例/学习成果展示
- 互动/福利帖（1条）：活动预告或学习tips
- 节点营销帖（视情况）：升阶季、开学季等

每条输出格式：
【发送时间】周X XX:XX
【内容类型】情感共鸣/产品价值/学员成果/互动福利
【正文】（150字内，口语化，带1-2个emoji）
【配图建议】
【私聊跟进话术】（针对点赞/评论的家长）
"""
    return prompt


def generate_schedule(region, stage, period, focus, custom_event, model="claude-sonnet-4-6"):
    content_lib = load_content_library()
    prompt = build_prompt(region, stage, period, focus, custom_event, content_lib)

    client = anthropic.Anthropic()
    message = client.messages.create(
        model=model,
        max_tokens=4096,
        messages=[{"role": "user", "content": prompt}],
    )

    return message.content[0].text


def save_output(content, region, period):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    date_str = datetime.now().strftime("%Y%m%d")
    region_label = "港澳台" if region == "hk" else "欧美澳"
    filename = f"{region_label}_{period}_{date_str}_排期.md"
    output_path = OUTPUT_DIR / filename
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(content)
    return output_path


def main():
    parser = argparse.ArgumentParser(description="LP企微朋友圈排期生成")
    parser.add_argument("--region", choices=["hk", "west"], required=True,
                        help="地区: hk=港澳台, west=欧美澳")
    parser.add_argument("--stage", default="all",
                        help="目标阶段: s1-s9 或 all")
    parser.add_argument("--period", choices=["week", "month"], default="week",
                        help="排期周期")
    parser.add_argument("--focus", choices=["upgrade", "retention", "renewal", "activity"],
                        default="retention", help="营销重点")
    parser.add_argument("--event", default=None, help="特殊节点，如'升阶季'")
    parser.add_argument("--model", default="claude-sonnet-4-6",
                        help="Claude模型ID")
    args = parser.parse_args()

    print(f"正在生成排期... 地区={args.region}, 阶段={args.stage}, "
          f"周期={args.period}, 重点={args.focus}")

    result = generate_schedule(
        region=args.region,
        stage=args.stage,
        period=args.period,
        focus=args.focus,
        custom_event=args.event,
        model=args.model,
    )

    output_path = save_output(result, args.region, args.period)
    print(f"排期已生成: {output_path}")


if __name__ == "__main__":
    main()

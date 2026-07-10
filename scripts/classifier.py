"""批量关键词分类器（Anthropic SDK 走 MiniMax 后端，走用户套餐）。

设计目标（一次性解决 4 个根本问题）：
- 错位防御：每批原子化（50 词 → 50 JSON），keyword + 数量严格校验，失败整批重试
- prompt 漂移防御：prompt 是 prompt_v1.txt 常量；temperature=0；model 固定
- 断点续跑：每批写 outputs/_checkpoints/batch_NNNN.json，启动时跳过已完成
- 对账：调用 build_excel + 输出对账报告

CLI：
  python scripts/classifier.py                          # 跑 inputs/keywords-*.txt 最新文件
  python scripts/classifier.py --input path.txt         # 指定输入
  python scripts/classifier.py --batch-size 50          # 自定义批大小
  python scripts/classifier.py --max-batches 2          # 只跑前 2 批（测试用）
  python scripts/classifier.py --reset                  # 清空 checkpoints 重跑
  python scripts/classifier.py --skip-excel             # 不调 build_excel（手动控制）
"""
import argparse
import functools
import json
import os
import re
import sys
import time
from pathlib import Path

import anthropic

# 强制 print 实时 flush（否则后台/管道场景下块缓冲，输出不可见）
print = functools.partial(print, flush=True)

# 让 scripts 包内模块能 import
sys.path.insert(0, str(Path(__file__).parent))

from parse_input import parse_keywords
from brands import load_brands, match_brand
from build_excel import build

# ── 常量（防漂移）────────────────────────────────────
PROMPT_PATH = Path(__file__).parent / "prompt_v1.txt"
BATCH_SIZE_DEFAULT = 50
MODEL = os.environ.get("ANTHROPIC_MODEL", "MiniMax-M3")
TEMPERATURE = 0  # 确定性，避免同 prompt 不同结果
MAX_TOKENS = 8000  # 50 词 × 6 字段 ≈ 3-4K tokens，留 buffer
MAX_RETRIES = 2  # 批级重试（PRD 3.3 节说 1 次，严格点 2 次）
CHECKPOINT_DIR = Path("outputs/_checkpoints")
CHECKPOINT_DIR.mkdir(parents=True, exist_ok=True)
MERGED_JSON = Path("outputs/_classifications.json")
INPUT_GLOB = "inputs/keywords-*.txt"


# ── 错位防御：严格校验 LLM 输出 ────────────────────────
def _validate_results(results, expected_keywords, batch_id):
    """严格校验 LLM 输出：数量匹配 + keyword 全部出现 + 字段齐全。

    返回 (ok, error_msg)。
    """
    if not isinstance(results, list):
        return False, f"批 {batch_id}: 输出不是数组（type={type(results).__name__}）"

    if len(results) != len(expected_keywords):
        return False, f"批 {batch_id}: 数量不匹配：输入 {len(expected_keywords)}，输出 {len(results)}"

    expected_lower = {kw.lower() for kw in expected_keywords}
    seen_lower = []
    for i, item in enumerate(results):
        if not isinstance(item, dict):
            return False, f"批 {batch_id} 第 {i+1} 项不是对象"
        kw = item.get("keyword")
        if not isinstance(kw, str):
            return False, f"批 {batch_id} 第 {i+1} 项 keyword 不是字符串"
        seen_lower.append(kw.lower())

    seen_set = set(seen_lower)
    missing = expected_lower - seen_set
    if missing:
        return False, f"批 {batch_id}: 缺失 keyword: {sorted(missing)[:3]}{'...' if len(missing)>3 else ''}"

    extra = seen_set - expected_lower
    if extra:
        return False, f"批 {batch_id}: 多出 keyword: {sorted(extra)[:3]}{'...' if len(extra)>3 else ''}"

    # 顺序一致性
    for i, (exp, got) in enumerate(zip(expected_keywords, seen_lower)):
        if exp.lower() != got:
            return False, f"批 {batch_id} 第 {i+1} 项顺序错位：期望 {exp!r}，实际 {got!r}"

    # 字段齐全性
    required_fields = {"keyword", "translation", "intent", "type", "monetization", "advice", "target_user"}
    for i, item in enumerate(results):
        missing_fields = required_fields - set(item.keys())
        if missing_fields:
            return False, f"批 {batch_id} 第 {i+1} 项缺字段: {missing_fields}"

    return True, ""


# ── 从 LLM 响应中提取 JSON 数组 ─────────────────────────
def _extract_json_array(text):
    """从 LLM 响应文本中提取 JSON 数组。

    处理 LLM 偶尔会包 markdown 代码块的情况（即使 prompt 禁止了）。
    """
    text = text.strip()
    # 去除 markdown 代码块包装
    fence_match = re.search(r"```(?:json)?\s*(\[.*?\])\s*```", text, re.DOTALL)
    if fence_match:
        text = fence_match.group(1)
    # 找到第一个 [ 到最后一个 ] 的范围
    start = text.find("[")
    end = text.rfind("]")
    if start == -1 or end == -1 or end <= start:
        raise ValueError("响应文本中找不到 JSON 数组")
    return json.loads(text[start:end + 1])


# ── 核心：单批分类（带重试） ─────────────────────────
def classify_batch(client, batch_keywords, batch_id):
    """分类一批关键词，含重试。返回 list[dict]（已通过校验）。"""
    prompt_template = PROMPT_PATH.read_text(encoding="utf-8")
    user_content = json.dumps(batch_keywords, ensure_ascii=False)

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = client.messages.create(
                model=MODEL,
                max_tokens=MAX_TOKENS,
                temperature=TEMPERATURE,
                messages=[
                    {"role": "user", "content": prompt_template},
                    {"role": "user", "content": f"输入关键词列表：\n{user_content}"},
                ],
            )
            text = response.content[0].text
            results = _extract_json_array(text)
            ok, err = _validate_results(results, batch_keywords, batch_id)
            if ok:
                return results, response.usage
            print(f"  ⚠ 批 {batch_id} 第 {attempt}/{MAX_RETRIES} 次校验失败: {err}")
        except Exception as e:
            print(f"  ⚠ 批 {batch_id} 第 {attempt}/{MAX_RETRIES} 次异常: {type(e).__name__}: {e}")
        if attempt < MAX_RETRIES:
            time.sleep(2)  # 简单退避

    raise RuntimeError(f"批 {batch_id} 重试 {MAX_RETRIES} 次仍失败")


# ── 品牌词 advice 覆盖值 ─────────────────────────────
# v1.5：让品牌词也走 LLM 分类（拿到真实 translation/type/target_user），
#       合并阶段只把 advice 字段改写成这个值，其他字段保留 LLM 的真实判断。
_BRAND_ADVICE = "跳过（品牌）"


def _apply_brand_advice(results, brand_set_lower):
    """对命中的品牌词，强制把 advice 覆盖为 '跳过（品牌）'。其他字段不动。

    brand_set_lower: set[str] —— 品牌关键词集合（小写），用于匹配 result.keyword。
    返回: int（被覆盖的条数）。
    """
    overwritten = 0
    for item in results:
        if item.get("keyword", "").lower() in brand_set_lower:
            if item.get("advice") != _BRAND_ADVICE:
                item["advice"] = _BRAND_ADVICE
                overwritten += 1
    return overwritten


# ── 变现路径归一化（v1.6）─────────────────────────────
# LLM 自由发挥会出现顺序/空格不一致，如：
#   "AdSense / 联盟" / "联盟 / AdSense" / "AdSense/联盟"
#   "AdSense / 联盟 / 电商" / "AdSense / 电商 / 联盟"
# 归一化策略：按 prompt 枚举的固定顺序排序 + 统一分隔符 " / "。
_MONETIZATION_ORDER = ["SaaS", "AdSense", "联盟", "电商", "导航", "待定"]
_MONETIZATION_RANK = {v: i for i, v in enumerate(_MONETIZATION_ORDER)}


def _canon_monetization(value: str) -> str:
    """归一化变现路径字符串。

    规则：
    1. 按 "/" 拆分
    2. 每项去前后空格；空字符串过滤掉
    3. 按 _MONETIZATION_RANK 排序（未知项排最后）
    4. 用统一分隔符 " / "（前后各一空格）连接

    例：
        "AdSense / 联盟"      → "AdSense / 联盟"
        "联盟 / AdSense"      → "AdSense / 联盟"
        "AdSense/联盟"        → "AdSense / 联盟"
        "AdSense / 联盟 / 电商" → "AdSense / 联盟 / 电商"
        ""                    → ""
    """
    if not value:
        return ""
    items = [s.strip() for s in value.split("/") if s.strip()]
    items.sort(key=lambda x: _MONETIZATION_RANK.get(x, 999))
    return " / ".join(items)


# ── 找最新输入文件 ──────────────────────────────────
def find_latest_input():
    candidates = sorted(Path("inputs").glob("keywords-*.txt"), reverse=True)
    if not candidates:
        raise FileNotFoundError(f"在 inputs/ 下找不到 keywords-*.txt")
    return candidates[0]


# ── 主流程 ──────────────────────────────────────
def run(input_path, batch_size=BATCH_SIZE_DEFAULT, max_batches=None, reset=False, skip_excel=False):
    if reset:
        for f in CHECKPOINT_DIR.glob("batch_*.json"):
            f.unlink()
        print(f"已清空 checkpoints，准备从头跑")

    # 1. 解析输入
    input_path = Path(input_path)
    print(f"[1/5] 解析输入: {input_path}")
    parsed = parse_keywords(input_path.read_text(encoding="utf-8"))
    keywords = [k for k, _ in parsed]
    print(f"  → {len(keywords)} 个关键词（去重后）")

    # 2. 品牌库匹配
    print(f"[2/5] 品牌库匹配")
    brands = load_brands("brands.txt")
    brand_hits = []  # [(keyword, brand), ...]   仅用于事后统计 + advice 覆盖
    to_classify = []
    for kw in keywords:
        hit = match_brand(kw, brands)
        if hit:
            brand_hits.append((kw, hit))
        # v1.5：品牌词也送 LLM（拿到真实 translation/type/target_user），
        #       合并阶段再把 advice 覆盖为 "跳过（品牌）"
        to_classify.append(kw)
    print(f"  → 命中 {len(brand_hits)} 个品牌词，全部 {len(to_classify)} 个进 LLM")

    # 3. 分批
    batches = [to_classify[i:i + batch_size] for i in range(0, len(to_classify), batch_size)]
    if max_batches is not None:
        batches = batches[:max_batches]
    total_batches = len(batches)
    print(f"[3/5] 分批: 批大小 {batch_size}，共 {total_batches} 批（将实际跑 {len(batches)} 批）")

    # 4. 逐批分类（断点续跑）
    print(f"[4/5] 开始分类")
    client = anthropic.Anthropic()  # 自动读 ANTHROPIC_BASE_URL + ANTHROPIC_AUTH_TOKEN
    total_input_tokens = 0
    total_output_tokens = 0
    failed_batches = []
    completed_batches = 0

    for batch_id, batch in enumerate(batches):
        ckpt = CHECKPOINT_DIR / f"batch_{batch_id:04d}.json"
        if ckpt.exists():
            print(f"  ✓ 批 {batch_id:04d} 已存在 checkpoint，跳过 ({len(batch)} 词)")
            completed_batches += 1
            continue

        print(f"  → 批 {batch_id:04d}/{total_batches - 1} 分类 {len(batch)} 词...")
        try:
            results, usage = classify_batch(client, batch, batch_id)
            # 原子写入：先写临时文件再 rename，避免半成品
            tmp = ckpt.with_suffix(".json.tmp")
            tmp.write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
            tmp.rename(ckpt)
            total_input_tokens += usage.input_tokens
            total_output_tokens += usage.output_tokens
            completed_batches += 1
            print(f"  ✓ 批 {batch_id:04d} 完成（in={usage.input_tokens} out={usage.output_tokens} tokens）")
        except Exception as e:
            print(f"  ✗ 批 {batch_id:04d} 失败: {e}")
            failed_batches.append((batch_id, batch, str(e)))
            # 不写 checkpoint，保留在 to_classify 中后续可重试

    if failed_batches:
        print(f"\n⚠ {len(failed_batches)} 批失败：{[(b, e[:50]) for b, _, e in failed_batches]}")
        print("  重新跑（不要 --reset）即可只重试这些批。")

    # 5. 合并 + 出 Excel
    print(f"\n[5/5] 合并结果")
    all_results = []
    for batch_id, batch in enumerate(batches):
        ckpt = CHECKPOINT_DIR / f"batch_{batch_id:04d}.json"
        if ckpt.exists():
            all_results.extend(json.loads(ckpt.read_text(encoding="utf-8")))

    # v1.5：品牌词 advice 覆盖（其他字段保留 LLM 真实判断）
    brand_kw_set = {kw.lower() for kw, _ in brand_hits}
    overwritten = _apply_brand_advice(all_results, brand_kw_set)
    print(f"  → 合并 {len(all_results)} 条；覆盖品牌 advice {overwritten} 条 → {_BRAND_ADVICE!r}")

    # v1.6：变现路径归一化（统一分隔符 + 顺序）
    canon_count = 0
    for item in all_results:
        old = item.get("monetization", "")
        new = _canon_monetization(old)
        if old != new:
            item["monetization"] = new
            canon_count += 1
    print(f"  → 变现路径归一化 {canon_count} 条")

    # 写入合并 JSON（喂 build_excel 用）
    MERGED_JSON.write_text(
        json.dumps(all_results, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"  → 写入 {MERGED_JSON}")

    # 对账报告
    expected_set = {kw.lower() for kw in keywords}
    got_set = {item["keyword"].lower() for item in all_results}
    missing = expected_set - got_set
    extra = got_set - expected_set
    print(f"\n=== 对账报告 ===")
    print(f"输入关键词: {len(keywords)}（去重后）")
    print(f"品牌命中:   {len(brand_hits)}")
    print(f"LLM 分类:   {len(all_results) - len(brand_hits)}")
    print(f"已合并:     {len(all_results)}")
    if missing:
        print(f"⚠ 缺失:     {len(missing)} 个（其中 {len(failed_batches)} 批失败导致）")
    else:
        print(f"✓ 无缺失")
    if extra:
        print(f"⚠ 多出:     {len(extra)} 个")
    print(f"\nToken 用量: 输入 {total_input_tokens:,} + 输出 {total_output_tokens:,} = {total_input_tokens + total_output_tokens:,}")

    if not skip_excel and not failed_batches:
        print(f"\n→ 出 Excel")
        try:
            output_dir = Path("outputs")
            output_dir.mkdir(parents=True, exist_ok=True)
            excel_path = build(MERGED_JSON, output_dir)
            print(f"✓ Excel: {excel_path}")
        except Exception as e:
            print(f"✗ Excel 生成失败: {e}")

    return {
        "total": len(keywords),
        "brand_hits": len(brand_hits),
        "llm_classified": len(all_results) - len(brand_hits),
        "failed_batches": failed_batches,
        "missing": missing,
    }


# ── CLI ──────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description="批量关键词分类器（走 MiniMax 套餐）")
    parser.add_argument("--input", type=str, help="输入文件路径（默认取 inputs/keywords-*.txt 最新）")
    parser.add_argument("--batch-size", type=int, default=BATCH_SIZE_DEFAULT, help=f"批大小（默认 {BATCH_SIZE_DEFAULT}）")
    parser.add_argument("--max-batches", type=int, default=None, help="最多跑几批（测试用）")
    parser.add_argument("--reset", action="store_true", help="清空 checkpoints 重跑")
    parser.add_argument("--skip-excel", action="store_true", help="不自动调 build_excel")
    args = parser.parse_args()

    input_path = Path(args.input) if args.input else find_latest_input()
    run(
        input_path=input_path,
        batch_size=args.batch_size,
        max_batches=args.max_batches,
        reset=args.reset,
        skip_excel=args.skip_excel,
    )


if __name__ == "__main__":
    main()

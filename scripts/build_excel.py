"""从分类结果 JSON 生成 Excel。

对话式入口的"出 Excel"步骤：
- 读 JSON（Claude 在对话里分类完的结果）
- 调用 url_builder 补全 4 列 URL
- 调用 generate_excel 出 xlsx

文件名遵循 PRD 4.6 节：关键词分类_YYYYMMDD_HHMM.xlsx，
同分钟内多次运行冲突时追加 _2、_3。
"""
import json
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from generate_excel import generate
from url_builder import build_all_urls


def build(json_path: str | Path, output_dir: str | Path = "outputs") -> Path:
    """读 JSON → 生成 Excel，返回输出路径。"""
    data = json.loads(Path(json_path).read_text(encoding="utf-8"))
    results = []
    for item in data:
        item = dict(item)
        item["urls"] = build_all_urls(item["keyword"])
        results.append(item)

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    base = output_dir / f"关键词分类_{timestamp}.xlsx"
    final_path = _resolve_collision(base)
    return generate(results, final_path)


def _resolve_collision(path: Path) -> Path:
    """同分钟内多次运行，文件名冲突时追加 _2、_3 ..."""
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    parent = path.parent
    for n in range(2, 1000):
        candidate = parent / f"{stem}_{n}{suffix}"
        if not candidate.exists():
            return candidate
    raise RuntimeError(f"无法找到可用文件名：{stem}{suffix}")


def summarize(path: str | Path) -> dict:
    """生成简要统计：v1.1 — 类型（工具/内容/带货/游戏/导航/空）+ 建议（建议/待研究/跳过）。"""
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    types = {"工具": 0, "内容": 0, "带货": 0, "游戏": 0, "导航": 0, "空": 0}
    advice = {"建议": 0, "待研究": 0, "跳过": 0}
    for item in data:
        t = item.get("type", "空")
        types[t] = types.get(t, 0) + 1
        a = item.get("advice", "")
        advice[a] = advice.get(a, 0) + 1
    return {"total": len(data), "type": types, "advice": advice}


if __name__ == "__main__":
    json_path = sys.argv[1] if len(sys.argv) > 1 else "outputs/_classifications.json"
    output_dir = sys.argv[2] if len(sys.argv) > 2 else "outputs"
    out = build(json_path, output_dir)
    stats = summarize(json_path)
    t = stats["type"]
    a = stats["advice"]
    print(f"已生成：{out}")
    print(f"共 {stats['total']} 词")
    print(f"建议={a.get('建议', 0)}、待研究={a.get('待研究', 0)}、跳过={a.get('跳过', 0)}")
    print(f"工具={t.get('工具', 0)}、内容={t.get('内容', 0)}、带货={t.get('带货', 0)}、游戏={t.get('游戏', 0)}、导航={t.get('导航', 0)}、空={t.get('空', 0)}")
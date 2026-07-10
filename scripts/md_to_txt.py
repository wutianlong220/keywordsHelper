"""md → txt 预处理：剥离 markdown 表格/列表标记。

跑关键词的预处理步骤：classifier.py 只接受 .txt 格式，
当 inputs/ 最新文件是 .md 时，Claude 调本脚本生成临时 .txt 再跑 classifier。

支持的 md 格式：
- 表格：`| col1 | col2 |` → 取 col1
- 表格分隔行：`| --- |` → 跳过
- 无序列表：`- keyword` 或 `* keyword` → 去掉标记
- 编号列表：`1. keyword` → 去掉标记
- 普通行：原样保留

CLI：
    python scripts/md_to_txt.py <input.md> <output.txt>
"""
import re
import sys
from pathlib import Path


# 表格分隔行：| --- | --- |、|---|---|、--- 都视为分隔行
_TABLE_SEP = re.compile(r"^\s*\|?[\s\-:|]+\|?\s*$")
# 无序列表前缀：- 或 *
_LIST_MARK = re.compile(r"^[\s]*[-\*]\s+")
# 编号列表前缀：1. 2. 等
_NUM_LIST = re.compile(r"^\s*\d+\.\s+")


def md_to_txt(md_text: str) -> str:
    """剥离 markdown 标记，返回纯文本（每行一个关键词）。"""
    out_lines = []
    for line in md_text.splitlines():
        stripped = line.strip()
        if not stripped:
            continue
        # 跳过表格分隔行（含 --- 或 ___ 或 :--）
        if _TABLE_SEP.match(stripped) and ("---" in stripped or "___" in stripped or ":--" in stripped):
            continue
        # 表格行：取第一列（第一个 | 之前的内容）
        if stripped.startswith("|"):
            cols = [c.strip() for c in stripped.split("|") if c.strip()]
            if cols:
                out_lines.append(cols[0])
            continue
        # 编号列表 / 无序列表：去掉前缀
        line = _NUM_LIST.sub("", stripped)
        line = _LIST_MARK.sub("", line)
        out_lines.append(line)
    # 末尾换行（POSIX 习惯）
    return "\n".join(out_lines) + "\n" if out_lines else ""


def main():
    if len(sys.argv) < 3:
        print("用法: python scripts/md_to_txt.py <input.md> <output.txt>", file=sys.stderr)
        sys.exit(1)
    md_path = Path(sys.argv[1])
    txt_path = Path(sys.argv[2])
    if not md_path.exists():
        print(f"✗ 输入不存在: {md_path}", file=sys.stderr)
        sys.exit(1)
    md_text = md_path.read_text(encoding="utf-8")
    converted = md_to_txt(md_text)
    txt_path.parent.mkdir(parents=True, exist_ok=True)
    txt_path.write_text(converted, encoding="utf-8")
    n_lines = len(converted.splitlines())
    print(f"✓ {md_path.name} → {txt_path.name}（{n_lines} 行）")


if __name__ == "__main__":
    main()
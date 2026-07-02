"""输入关键词解析。

实现 PRD 第 4.5 节：
- 一行一个关键词
- 括号备注（中文（）或英文()内）剥离
- 空行忽略
- 重复关键词去重（保留首次出现位置）

返回 [(keyword, note), ...]，note 为空串表示无备注。

限制：MVP 不支持嵌套括号（用户备注通常也不会嵌套）。
"""
import re
from pathlib import Path


# 匹配括号及其内部内容（含括号本身），同时覆盖中文（）和英文 ()
# 非贪婪匹配，剥掉最内层括号对（嵌套场景会从内向外剥多次）
_NOTE_PATTERN = re.compile(r"[（(][^）)]*[）)]")


def parse_keywords(text: str) -> list[tuple[str, str]]:
    """解析关键词文本。

    返回 [(keyword, note), ...]：
    - keyword: 剥离括号备注后的关键词（已 strip）
    - note: 括号内文字；无备注则为空串

    规则：
    - 一行一个
    - 空行忽略
    - 重复关键词去重（按 keyword 大小写不敏感比较，保留首次）
    """
    seen: set[str] = set()
    results: list[tuple[str, str]] = []

    for line in text.splitlines():
        # 先抓括号备注（要剥离前先抓）
        note_match = _NOTE_PATTERN.search(line)
        note = note_match.group(0)[1:-1].strip() if note_match else ""

        # 剥离括号及其内容
        cleaned = _NOTE_PATTERN.sub("", line).strip()
        if not cleaned:
            continue

        # 去重（按 keyword 大小写不敏感）
        key = cleaned.lower()
        if key in seen:
            continue
        seen.add(key)
        results.append((cleaned, note))

    return results


def parse_keywords_file(path: str | Path) -> list[tuple[str, str]]:
    """从文件读取并解析。"""
    return parse_keywords(Path(path).read_text(encoding="utf-8"))
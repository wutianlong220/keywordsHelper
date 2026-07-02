"""brands.txt 解析与子串匹配。

实现 PRD 第 2.2 节 Story 3 的品牌库要求：
- 子串包含匹配（大小写不敏感）
- 注释行（# 开头）和空行忽略
- 命中后由调用方在 note 字段追加 [品牌库:xxx] 溯源
"""
from pathlib import Path


def load_brands(path: str | Path) -> list[str]:
    """读取 brands.txt，返回小写品牌词的有序列表。

    规则：
    - 跳过空行
    - 跳过以 # 开头的注释行
    - 保留原始顺序
    - 全部转小写（匹配时也转小写，实现大小写不敏感）
    """
    brands: list[str] = []
    for line in Path(path).read_text(encoding="utf-8").splitlines():
        s = line.strip()
        if not s or s.startswith("#"):
            continue
        brands.append(s.lower())
    return brands


def match_brand(keyword: str, brands: list[str]) -> str | None:
    """子串包含匹配（大小写不敏感）。

    返回第一个命中的品牌（统一小写），否则 None。
    遍历顺序按 brands 列表顺序，因此列表前面的品牌优先。
    brands 列表元素不要求预小写 —— 函数内部会转小写。
    """
    kw = keyword.lower()
    for brand in brands:
        if brand.lower() in kw:
            return brand.lower()
    return None
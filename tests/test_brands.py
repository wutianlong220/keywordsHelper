"""brands.py 的单元测试。"""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent.parent / "scripts"))

from brands import load_brands, match_brand


# ──────────────── load_brands ────────────────

class TestLoadBrands:
    def test_基本加载(self, tmp_path: Path):
        f = tmp_path / "brands.txt"
        f.write_text("nike\nadidas\npuma\n", encoding="utf-8")
        assert load_brands(f) == ["nike", "adidas", "puma"]

    def test_跳过注释行(self, tmp_path: Path):
        f = tmp_path / "brands.txt"
        f.write_text("# 运动品牌\nnike\n# 餐饮品牌\nstarbucks\n", encoding="utf-8")
        assert load_brands(f) == ["nike", "starbucks"]

    def test_跳过空行(self, tmp_path: Path):
        f = tmp_path / "brands.txt"
        f.write_text("\nnike\n\n\nstarbucks\n\n", encoding="utf-8")
        assert load_brands(f) == ["nike", "starbucks"]

    def test_转小写(self, tmp_path: Path):
        f = tmp_path / "brands.txt"
        f.write_text("Nike\nADIDAS\nPuma\n", encoding="utf-8")
        assert load_brands(f) == ["nike", "adidas", "puma"]

    def test_保留原始顺序(self, tmp_path: Path):
        f = tmp_path / "brands.txt"
        f.write_text("starbucks\nnike\namazon\n", encoding="utf-8")
        assert load_brands(f) == ["starbucks", "nike", "amazon"]

    def test_首尾空白(self, tmp_path: Path):
        f = tmp_path / "brands.txt"
        f.write_text("  nike  \n\tadidas\t\n", encoding="utf-8")
        # 实际 brands.txt 中"under armour"等带空格的应该整体 trim
        assert load_brands(f) == ["nike", "adidas"]

    def test_空文件(self, tmp_path: Path):
        f = tmp_path / "brands.txt"
        f.write_text("", encoding="utf-8")
        assert load_brands(f) == []


# ──────────────── match_brand ────────────────

class TestMatchBrand:
    def test_命中_简单品牌(self):
        assert match_brand("nike shoes", ["nike"]) == "nike"

    def test_命中_大小写不敏感(self):
        assert match_brand("NIKE SHOES", ["nike"]) == "nike"
        assert match_brand("nike shoes", ["NIKE"]) == "nike"

    def test_命中_关键词在中间(self):
        assert match_brand("buy nike today", ["nike"]) == "nike"

    def test_不命中_无品牌(self):
        assert match_brand("zerogpt", ["nike", "adidas"]) is None

    def test_不命中_子串误伤示例(self):
        # PRD 决策记录说"删除 apple 入库"是因为会误伤 pineapple
        # 这里验证 apple 不在库时不会误伤
        assert match_brand("pineapple", ["nike", "amazon"]) is None

    def test_多品牌_返回第一个(self):
        # brands 列表中前面的优先
        result = match_brand("nike adidas", ["nike", "adidas"])
        assert result == "nike"

    def test_命中_多词品牌(self):
        # "under armour" 是两个词的组合
        assert match_brand("under armour shirt", ["under armour"]) == "under armour"

    def test_不命中_部分匹配也算(self):
        # 子串包含匹配，不是整词匹配
        # 例：库里有 "kfc" → "kfc menu" 命中
        assert match_brand("kfc menu", ["kfc"]) == "kfc"

    def test_空品牌库(self):
        assert match_brand("nike", []) is None
"""parse_input.py 的单元测试。"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "scripts"))

from parse_input import parse_keywords, parse_keywords_file


class TestParseKeywords:
    def test_基本单行(self):
        assert parse_keywords("zerogpt") == [("zerogpt", "")]

    def test_多行(self):
        text = "zerogpt\nthunderstorm tracker uk\npag ibig loan tracker"
        assert parse_keywords(text) == [
            ("zerogpt", ""),
            ("thunderstorm tracker uk", ""),
            ("pag ibig loan tracker", ""),
        ]

    def test_跳过空行(self):
        text = "zerogpt\n\n\nthunderstorm\n"
        assert parse_keywords(text) == [("zerogpt", ""), ("thunderstorm", "")]

    def test_中文括号备注剥离(self):
        text = "jmail.world （备注：推特热搜，爱泼斯坦案档案站）"
        assert parse_keywords(text) == [("jmail.world", "备注：推特热搜，爱泼斯坦案档案站")]

    def test_英文括号备注剥离(self):
        text = "example (note here)"
        assert parse_keywords(text) == [("example", "note here")]

    def test_中英文括号混用(self):
        text = "foo (英文备注)\nbar （中文备注）"
        assert parse_keywords(text) == [
            ("foo", "英文备注"),
            ("bar", "中文备注"),
        ]

    def test_去重_完全相同(self):
        text = "zerogpt\nzerogpt\nthunderstorm"
        assert parse_keywords(text) == [("zerogpt", ""), ("thunderstorm", "")]

    def test_去重_大小写不敏感(self):
        text = "Zerogpt\nzerogpt\nZEROGPT"
        assert parse_keywords(text) == [("Zerogpt", "")]

    def test_去重_保留首次(self):
        text = "zerogpt\nnike shoes\nzerogpt"
        assert parse_keywords(text) == [("zerogpt", ""), ("nike shoes", "")]

    def test_备注里有括号_MVP不支持(self):
        # MVP 不支持嵌套括号 —— 当前 regex 会从最内层开始剥离
        # 实际效果：example (note with (inner) parens) → 剥掉 (inner) 后
        # 剩下 "example (note with  parens)" 仍含外层括号
        # 用户应避免在备注中嵌套括号
        text = "example (note with (inner) parens)"
        result = parse_keywords(text)
        # 不崩、keyword 非空即可，不强断言具体形态
        assert len(result) == 1
        assert result[0][0]  # 非空

    def test_空白行与全空白(self):
        text = "  \nzerogpt\n   \n"
        assert parse_keywords(text) == [("zerogpt", "")]

    def test_关键词首尾空白(self):
        text = "  zerogpt  \n  thunderstorm  "
        # strip 后再处理
        assert parse_keywords(text) == [
            ("zerogpt", ""),
            ("thunderstorm", ""),
        ]

    def test_空输入(self):
        assert parse_keywords("") == []

    def test_纯备注无关键词(self):
        # 只有括号内容（如 "(just a note)"），剥离后为空，被忽略
        text = "(just a note)"
        assert parse_keywords(text) == []


class TestParseKeywordsFile:
    def test_从文件读取(self, tmp_path: Path):
        f = tmp_path / "kw.txt"
        f.write_text("zerogpt\nnike shoes （鞋）\n", encoding="utf-8")
        assert parse_keywords_file(f) == [("zerogpt", ""), ("nike shoes", "鞋")]
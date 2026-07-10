"""url_builder 的单元测试。"""
import sys
from pathlib import Path

# 把 scripts/ 加入 import 路径
sys.path.insert(0, str(Path(__file__).parent.parent / "scripts"))

from url_builder import (
    build_intitle_url,
    build_serp_url,
    build_trends_url,
    build_ahrefs_url,
    build_all_urls,
)


class TestBuildIntitleUrl:
    def test_普通关键词(self):
        url = build_intitle_url("zerogpt")
        # 模板中 %22 是硬编码双引号，zerogpt 不需要编码
        assert url == "https://www.google.com/search?q=intitle:%22zerogpt%22&hl=en&gl=us"

    def test_带空格的关键词(self):
        url = build_intitle_url("thunderstorm tracker uk")
        # 空格 → %20
        assert url == "https://www.google.com/search?q=intitle:%22thunderstorm%20tracker%20uk%22&hl=en&gl=us"

    def test_带连字符(self):
        url = build_intitle_url("cafe-au-lait")
        # 连字符在 RFC 3986 unreserved 集合中，不编码
        assert url == "https://www.google.com/search?q=intitle:%22cafe-au-lait%22&hl=en&gl=us"

    def test_带特殊字符(self):
        # & 在 URL 中是分隔符，必须编码
        url = build_intitle_url("q&a")
        assert url == "https://www.google.com/search?q=intitle:%22q%26a%22&hl=en&gl=us"

    def test_大小写保留(self):
        url = build_intitle_url("NIKE")
        assert url == "https://www.google.com/search?q=intitle:%22NIKE%22&hl=en&gl=us"


class TestBuildSerpUrl:
    def test_普通关键词(self):
        url = build_serp_url("zerogpt")
        assert url == "https://www.google.com/search?q=zerogpt&hl=en&gl=us"

    def test_带空格(self):
        url = build_serp_url("pag ibig loan tracker")
        assert url == "https://www.google.com/search?q=pag%20ibig%20loan%20tracker&hl=en&gl=us"


class TestBuildTrendsUrl:
    def test_普通关键词(self):
        url = build_trends_url("zerogpt")
        # 模板中 date=today%201-m 的 %20 是硬编码，保留；不带 geo 参数即默认全球
        assert url == "https://trends.google.com/trends/explore?q=zerogpt&date=today%201-m"

    def test_带空格(self):
        url = build_trends_url("thunderstorm tracker uk")
        assert url == "https://trends.google.com/trends/explore?q=thunderstorm%20tracker%20uk&date=today%201-m"


class TestBuildAhrefsUrl:
    def test_普通关键词(self):
        url = build_ahrefs_url("zerogpt")
        assert url == "https://ahrefs.com/keyword-difficulty/?country=us&input=zerogpt"

    def test_带空格(self):
        url = build_ahrefs_url("pag ibig loan tracker")
        assert url == "https://ahrefs.com/keyword-difficulty/?country=us&input=pag%20ibig%20loan%20tracker"


class TestBuildAllUrls:
    def test_返回4个键(self):
        urls = build_all_urls("zerogpt")
        assert set(urls.keys()) == {"intitle", "serp", "trends", "ahrefs"}

    def test_4个URL_前缀正确(self):
        urls = build_all_urls("zerogpt")
        assert urls["intitle"].startswith("https://www.google.com/search?q=intitle:")
        assert urls["serp"].startswith("https://www.google.com/search?q=zerogpt")
        assert urls["trends"].startswith("https://trends.google.com/trends/explore?q=zerogpt")
        assert urls["ahrefs"].startswith("https://ahrefs.com/keyword-difficulty/?country=us&input=zerogpt")
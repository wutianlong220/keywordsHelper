"""URL 模板拼接与编码。

实现 PRD 第 4.3 节定义的 4 个 URL 模板。
编码策略：模板中的 %22、%201 等固定部分保留原样，仅对 {kw} 占位符做 URL 编码。
"""
from urllib.parse import quote


# PRD 4.3 节规定的 URL 模板
INTITLE_URL = "https://www.google.com/search?q=intitle:%22{kw}%22&hl=en&gl=us"
SERP_URL = "https://www.google.com/search?q={kw}&hl=en&gl=us"
TRENDS_URL = "https://trends.google.com/trends/explore?q={kw}&geo=US&date=today%201-m"
AHREFS_URL = "https://ahrefs.com/keyword-difficulty/?country=us&input={kw}"


def build_intitle_url(keyword: str) -> str:
    """intitle 检查 URL。模板中 %22 是硬编码的双引号，仅编码关键词部分。"""
    return INTITLE_URL.replace("{kw}", quote(keyword, safe=""))


def build_serp_url(keyword: str) -> str:
    """SERP 检查 URL。"""
    return SERP_URL.replace("{kw}", quote(keyword, safe=""))


def build_trends_url(keyword: str) -> str:
    """谷歌趋势 URL（30 天）。模板中 %201-m 段保留，编码关键词。"""
    return TRENDS_URL.replace("{kw}", quote(keyword, safe=""))


def build_ahrefs_url(keyword: str) -> str:
    """Ahrefs KD URL。"""
    return AHREFS_URL.replace("{kw}", quote(keyword, safe=""))


def build_all_urls(keyword: str) -> dict:
    """一次性返回 4 个 URL。"""
    return {
        "intitle": build_intitle_url(keyword),
        "serp": build_serp_url(keyword),
        "trends": build_trends_url(keyword),
        "ahrefs": build_ahrefs_url(keyword),
    }
# keywordsHelper

英文关键词批量分类工具 —— 对话式调用 LLM 完成语义分类，输出带可点击链接与颜色标注的 Excel 表。

## 核心特性

- **双轴分类**：每个关键词分两个独立维度
  - **类型**（关键词天然适合的网站类型）：工具 / 内容 / 带货 / 游戏 / 导航 / 空
  - **建议**（给用户的行动建议）：建议 / 待研究 / 跳过
- **品牌库过滤**：`brands.txt` 子串匹配命中 → 自动归为「跳过」，并加 `[品牌库:xxx]` 溯源标签
- **可点击链接**：每行附 4 个深度检查链接（intitle / SERP / 谷歌趋势 / Ahrefs KD）
- **Excel 格式**：冻结首行 + 筛选器 + 行级/列级着色 + 收藏列（深绿背景 + 文字 toggle）

## 目录结构

```
keywordsHelper/
├─ PRD.md                     # 产品需求文档（详细规范 + 决策记录）
├─ 关键词筛选完整流程.md         # 原始业务流程
├─ 待优化指南.md               # 未来功能备忘
├─ brands.txt                 # 品牌黑名单（手工维护，一行一个）
├─ scripts/                   # Python 源码
│  ├─ generate_excel.py       # 接收分类结果 JSON，生成 xlsx
│  ├─ build_excel.py          # 读 JSON → 补 URL → 出 Excel
│  ├─ url_builder.py          # URL 模板拼接与编码
│  └─ parse_input.py          # 输入关键词解析
├─ tests/                     # pytest 测试套件
├─ inputs/                    # 用户输入关键词（私有，不入 git）
└─ outputs/                   # Excel 输出（私有，不入 git）
```

## 快速开始

```bash
# 1. 安装依赖
python3 -m venv venv
source venv/bin/activate
pip install openpyxl pytest

# 2. 准备关键词（两种方式任选）
#    a) 放到 inputs/keywords-YYYY-MM-DD.txt（一行一个）
#    b) 直接在 Claude Code 对话里贴

# 3. 在 Claude Code 里说："处理 inputs/keywords-XXX.txt"
#    → 自动：品牌库匹配 → LLM 分类 → 生成 Excel → 返回路径 + 统计

# 4. 跑测试
venv/bin/python -m pytest tests/ -v
```

## 品牌库维护

编辑 `brands.txt`：

- 一行一个品牌关键词
- `#` 开头为注释行
- 空行忽略
- 子串匹配（大小写不敏感）—— 例：库里有 `nike` → `buy nike shoes` 自动归「跳过」
- **注意**：子串匹配会误伤（如 `apple` 命中 `pineapple`），加短词时谨慎

## 隐私说明

为避免泄露你的关键词研究内容，以下文件/目录**不入 git**（但保留目录本身）：

- `inputs/*` —— 你正在研究的真实关键词
- `outputs/*` —— 历史 Excel 分类结果 + 中间 JSON

如需在仓库里演示效果，可自行在 `inputs/` 放脱敏示例 + 跑一次分类（结果会出现在本地 `outputs/`，git 看不到）。

## 详细规范

产品需求、字段定义、URL 模板、决策记录见 [PRD.md](./PRD.md)。
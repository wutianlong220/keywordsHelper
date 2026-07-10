# 关键词分类项目约定

## 用户偏好
- 用户希望用**自然语言**指挥 Claude 跑全流程，不自己敲命令。

## 触发词
**「跑关键词」** —— 用户说「跑关键词」「帮我跑关键词」「跑一下关键词」等都触发。

## 触发后自动执行

### Step 1: 选输入文件
- 默认取 `inputs/keywords-*.txt` 最新
- 如果最新是 `.md`（不是 .txt）→ 先调 `scripts/md_to_txt.py` 转成临时 txt（`inputs/_tmp_keywords.txt`），用 `--input` 喂给 classifier，跑完删临时文件
- 用户指定文件名 → 用 `--input <路径>`

### Step 2: 跑分类（默认重置 + 全量）
```bash
cd /Users/xingchen/GitHub/keywordsHelper && venv/bin/python scripts/classifier.py --reset
```
- 临时 txt 场景：`venv/bin/python scripts/classifier.py --reset --input inputs/_tmp_keywords.txt`
- 跑完删除 `inputs/_tmp_keywords.txt`

### Step 3: 实时反馈
- 每批完成后通知用户（不等全部跑完再统一汇报）

### Step 4: 完成后
- 告知 Excel 完整路径（`outputs/关键词分类_YYYYMMDD_HHMM.xlsx`）

## 输入输出约定
- 输入：`inputs/keywords-*.txt` 或 `inputs/*.md`（默认取最新；用户指定文件名时用 `--input`）
- 输出：`outputs/关键词分类_YYYYMMDD_HHMM.xlsx`
- 跑批参数：批大小 50、温度 0、模型 `MiniMax-M3`、max_tokens 8000

## md → txt 转换（预处理）
- 入口：`scripts/md_to_txt.py <input.md> <output.txt>`
- 支持：表格行取首列 / 无序列表剥前缀 / 编号列表剥前缀 / 表格分隔行跳过
- 临时文件：`inputs/_tmp_keywords.txt`（跑完删）
- **不做 U+FFFD 清洗**（那是 classifier 的 parse_input.py 职责）

## 变体指令
- "跑前 N 批测试" → 加 `--max-batches N`
- "接着上次继续跑" → 不加 `--reset`
- "只分类，不出 Excel" → 加 `--skip-excel`

## 异常处理
- 单批失败：如实告知（已知原因如 `?` 字符词等），跑完后再单独补
- LLM 调用失败：报告错误，让用户决定（不擅自重试过多）
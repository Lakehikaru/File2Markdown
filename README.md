# file-to-markdown

一个 Claude Code Skill，通过 [markdown.new](https://markdown.new) API 将文件转换为 Markdown。支持 PDF、DOCX、XLSX、图片等 20+ 种格式，自动拆分超过 10MB 的大文件。

## 快速上手

```bash
# 1. 克隆仓库
git clone https://github.com/Lakehikaru/File2Markdown.git
cd File2Markdown

# 2. 安装依赖
pip install requests pypdfium2 openpyxl

# 3. 直接使用（无需 API key）
python .claude/scripts/file2md.py your_file.pdf

# 或在 Claude Code 中打开此目录，直接说：
# "把 report.pdf 转成 markdown"
```

## 功能特性

- 调用 markdown.new 免费 API，无需注册，每天 500 次
- 大文件自动拆分：PDF 按页、XLSX 按 sheet、DOCX 按段落
- 拆分后逐片上传转换，最终合并为完整 Markdown
- 支持批量转换多个文件
- 纯 Python 实现，依赖少

## 安装

### 1. 安装 Python 依赖

```bash
pip install requests pypdfium2 openpyxl
```

### 2. 集成到 Claude Code

**方式 A：克隆仓库（推荐团队使用）**

```bash
git clone https://github.com/Lakehikaru/File2Markdown.git
cd File2Markdown
# Claude Code 会自动识别 .claude/skills/ 下的 skill
```

**方式 B：复制到全局目录（个人使用）**

```bash
# 复制 skill 定义
cp .claude/skills/file-to-markdown.md ~/.claude/skills/

# 复制脚本
mkdir -p ~/.claude/scripts
cp .claude/scripts/file2md.py ~/.claude/scripts/
```

使用方式 B 时，需要将 skill 文件中的脚本路径改为绝对路径：
```
python ~/.claude/scripts/file2md.py
```

## 使用方法

### 通过 Claude Code（推荐）

在 Claude Code 中直接说：

> "把 report.pdf 转成 markdown"
>
> "把这个目录下所有 PDF 转成 md"

Claude 会自动调用 file-to-markdown skill 完成转换。

### 命令行直接使用

```bash
# 单个文件
python .claude/scripts/file2md.py document.pdf

# 指定输出路径
python .claude/scripts/file2md.py document.pdf -o output.md

# 批量转换
python .claude/scripts/file2md.py file1.pdf file2.docx file3.xlsx

# 使用 API key（可选，提高速率限制）
python .claude/scripts/file2md.py document.pdf --api-key mk_xxxxx
```

## 支持格式

| 格式 | 直接转换 | 大文件自动拆分 |
|------|---------|--------------|
| PDF | ✓ | ✓ 按页拆分 |
| DOCX | ✓ | ✓ 按段落拆分 |
| XLSX | ✓ | ✓ 按 sheet 拆分 |
| 图片 (JPG/PNG/WebP) | ✓ | - |
| HTML/XML/CSV/JSON | ✓ | - |
| ODS/ODT/Numbers | ✓ | - |

## 项目结构

```
.claude/
├── skills/
│   └── file-to-markdown.md   # Skill 定义（Claude Code 读取）
└── scripts/
    └── file2md.py             # 转换脚本
```

## 限制

- markdown.new API 单次上传限制 10MB
- 免费额度：500 次/天/IP
- 需要网络连接

## License

MIT

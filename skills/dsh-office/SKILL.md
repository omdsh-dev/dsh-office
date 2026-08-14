---
name: dsh-office
description: 生成、读取、编辑 Office 文档（xlsx 电子表格 / PDF / pptx 演示文稿）——原生渲染，中文友好，适合报告、合同、数据表、演示文稿等交付场景
triggers: [excel, xlsx, 表格, 电子表格, spreadsheet, pdf, 报告, 合同, 文档, export, pptx, ppt, 演示, 幻灯片, presentation, slide]
---

# dsh-office（Office 文档工具）

@huiliyi37/dsh-office 为 DeepSeek Harness 提供 7 个 Office 文档工具。参考 anthropics/skills（Apache 2.0）提炼。

## 工具总览

| 工具 | 用途 |
|------|------|
| `xlsx_write` | 二维数组 → 新 `.xlsx`（公式单元格、表头加粗、列宽、数字格式） |
| `xlsx_read` | 列出工作表 / 读指定表为 markdown 表格（range 分页、max_rows 截断带续读提示） |
| `xlsx_edit` | 增表、改单元格（值或公式）、追加行 |
| `pdf_create` | 内容块数组 → PDF（标题/段落/表格/列表/代码块，CJK 自动字体，页码可选） |
| `pdf_read` | 按页提取 PDF 文本（`--- Page N ---` 标记，start_page/end_page 分页） |
| `pptx_create` | 幻灯片定义 → `.pptx`（7 种版式 + 主题 + 演讲者备注） |
| `pptx_read` | 幻灯片文本 → markdown（可选含备注） |

## 何时使用

- 用户要**可交付文件**（.xlsx/.pdf/.pptx），不是 Markdown 就够的场景
- 数据适合表格/图表呈现，或需要正式排版（报告、合同、演示）
- 需要从已有 Office 文件中提取内容进上下文（读回自查、审查导出结果）

纯文本/Markdown 能交付时不用——更轻、可 diff。

## 大文件读取纪律（重要）

- `xlsx_read` 单次最多 500 行（默认 200），超出会截断并在末尾给出
  `Continue with range_start: "A{n}"`——**照提示继续读下一页**，不要猜测范围；
  需要更多行时显式传 `max_rows`（如 500）
- `pdf_read` 输出带页码标记且 8000 字符截断，超长时按
  `Continue with start_page: {n}` 提示继续；只想看某几页时直接传 `start_page`/`end_page`
- 不要为"读完整个文件"重复调用——按需读，读完相关部分就停

## 生成纪律

1. **结构先行**：数据用表格/列表呈现，不写流水句；PPT 每页一个主题
2. **生成后自查**：用 `xlsx_read`/`pdf_read`/`pptx_read` 读回产出，确认内容完整、无占位符（lorem / xxx / TODO）、中文渲染正常
3. **PDF 中文**：系统无 CJK 字体时会显式警告——告知用户，建议装 Noto Sans CJK 或改用英文
4. **正式交付**：PDF 默认带 `page_numbers: true`；PPT 大演示加演讲者备注（`notes` 字段）

## 组合示例

- 周报：`xlsx_write` 数据表 → `pdf_create` 排版成报告（表格块直接引用数据）
- 演示：`pptx_create` 生成 → `pptx_read` 自查 → 修正再生成
- 合同/长文：`pdf_create` 生成 → `pdf_read` 抽查关键页

# @huiliyi37/dsh-office

面向 [DeepSeek Harness](https://github.com/deepseek-ai/deepseek-harness)（`dsh`）的 Office 文档工具插件：生成、读取、编辑电子表格（`.xlsx`）、PDF 和演示文稿（`.pptx`）。

移植自 [天枢](https://github.com/Tianshu-Tui) 终端编程智能体的 office 插件（上游同为 Apache-2.0 许可），适配 dsh 的 cordis 工具模型。

## 工具列表

| 工具 | 功能 |
| --- | --- |
| `xlsx_read` | 列出 `.xlsx` 工作表，或将指定工作表读取为 markdown 表格（大文件支持 range 分页；公式单元格保留公式文本） |
| `xlsx_write` | 将二维数组写入新的 `.xlsx`（支持公式单元格、表头加粗、列宽、数字格式） |
| `xlsx_edit` | 编辑已有 `.xlsx`：新增工作表、更新单元格（值或公式）、追加行 |
| `pdf_create` | 生成真实 PDF：标题、段落、表格、列表、代码块、页脚页码；中文内容通过自动探测的系统字体渲染 |
| `pdf_read` | 提取 PDF 文本供上下文阅读 |
| `pptx_create` | 按幻灯片定义生成 `.pptx`（标题 / 章节 / 内容 / 双栏 / 图片 / 表格 / 图表），支持主题配色与演讲者备注 |
| `pptx_read` | 将幻灯片文本提取为 markdown，可选包含演讲者备注 |

## 安装

使用官方插件命令安装到 dsh profile（本包声明了 `dsh.bundle` manifest，`dsh plugin` 会将其作为 bundle 层激活）：

```sh
dsh plugin --profile <名称> add @huiliyi37/dsh-office
dsh --profile <名称>
```

首次执行 `dsh plugin` 会初始化 profile（`@deepseek-ai/dsh-base` 作为首个 bundle）并把本包追加到 profile 的 `bundles` 列表。之后启动该 profile 即自动加载插件并注册全部七个工具。

手动安装备选——把包装到 Node 可解析的位置，并在自己的 `cordis.patch.yml` 中引用：

```sh
npm install @huiliyi37/dsh-office
```

```yaml
# cordis.patch.yml
- insert:
    - id: dsh-office
      name: '@huiliyi37/dsh-office'
```

## 使用示例

```jsonc
// xlsx_write — 创建工作簿
{ "file_path": "report.xlsx", "data": [["Name", "Score"], ["Alice", 92]], "header_bold": true }

// pdf_create — 带标题、表格和列表的文档
{
  "destination_path": "doc.pdf",
  "title": "Quarterly Report",
  "content": [
    { "type": "heading", "text": "Summary" },
    { "type": "table", "headers": ["Region", "Revenue"], "rows": [["APAC", "120"]] },
    { "type": "list", "items": ["Alpha", "Beta"] }
  ],
  "page_numbers": true
}

// pptx_create — 标题页 + 要点页的演示文稿
{
  "destination_path": "deck.pptx",
  "slides": [
    { "type": "title", "title": "Roadmap 2026" },
    { "type": "content", "title": "Highlights", "items": ["Plugin runtime", "Office tools"] }
  ]
}
```

## Skill（使用指导）

本包附带使用指导 skill（`skills/SKILL.md`，anthropics 兼容格式），教会模型大文件分页读取与生成纪律。安装到 skill 发现根目录：

```sh
mkdir -p ~/.dsh/skills && cp -r node_modules/@huiliyi37/dsh-office/skills/dsh-office ~/.dsh/skills/
```

## 开发

```sh
npm install
npm run build   # tsc → lib/
npm test        # vitest：经工具执行路径的往返测试
```

## 许可

Apache License 2.0。工具逻辑移植自天枢 office 插件（同为 Apache-2.0 许可，版权归 Tianshu contributors）；各文件头保留来源说明。

# pptx_cli

`pptx_cli` 是一个面向 Agent 的 PPT 生产命令行工具。首版目标不是让 Agent 直接操作复杂的 PPT 对象模型，而是围绕“模板 + JSON 表单 + 生成/预览”建立稳定工作流。

## 当前状态

当前已完成任务 001、003、004、005 和 006 的首版能力：

- 建立了可运行的 CLI 骨架
- 明确了首版技术选型与目录结构
- 提供了一个最小 demo 命令
- 补充了 PPT 解析、写入、预览的技术路线说明
- 支持为指定 slide 生成候选对象 JSON
- 支持通过 LibreOffice 渲染高还原度预览图
- 支持输出目标检测风格的标注框预览图
- 支持 `template create --from -> template show/add_slide/save -> edit create/fill/save` 的 Agent 友好工作流
- 支持通过后台 server 复用当前 PPT 的加载与解析现场
- 支持创建模板草稿 JSON，并逐页确认模板字段
- 支持生成 `template.pptx + manifest.json` 模板包
- 支持通过命令行 `--slide + --field` 填充模板并生成新的 PPTX
- 支持通过 `edit fill_template --content '<json>'` 追加嵌套 Container 主体内容

## 技术选型

- **语言**：Python 3.11+
- **依赖管理**：`uv` + `pyproject.toml` + 本地 `.venv`
- **CLI 框架**：标准库 `argparse`
- **PPT 读写路线**：首选 `python-pptx`
- **数据模型路线**：内部采用 Python `dataclass`；对外 JSON 协议采用显式 schema 文档，后续可升级到 Pydantic
- **预览路线**：通过 LibreOffice headless 将 PPT 转 PDF/图片

更完整的论证见 [docs/tech-selection.md](docs/tech-selection.md)。会话模式说明见 [docs/session-lifecycle.md](docs/session-lifecycle.md)。

## 目录结构

```text
.
├── docs/
│   └── tech-selection.md
├── examples/
│   └── demo-form.json
├── pptxcli
├── pyproject.toml
├── uv.lock
├── src/
│   └── pptx_cli/
│       ├── __init__.py
│       ├── __main__.py
│       ├── cli.py
│       ├── inspect.py
│       ├── models.py
│       ├── show.py
│       └── template_ops.py
├── tasks/
└── tests/
    ├── test_cli.py
│   ├── test_show.py
│   └── test_template.py
```

## 快速开始

### 直接运行仓库内脚本

```bash
uv venv .venv
uv sync
./pptxcli --help
./pptxcli demo form
```

### 作为 Python 包运行

```bash
PYTHONPATH=src python3 -m pptx_cli --help
```

### 使用 uv 运行项目命令

```bash
uv run pptxcli --help
uv run pptxcli demo form
```

## 当前命令

- `pptxcli --help`：查看帮助
- `pptxcli version`：查看版本
- `pptxcli demo form`：输出最小表单 JSON 示例
- `pptxcli tech`：输出技术路线摘要
- `pptxcli inspect --input ./demo.pptx --slide 0`：输出指定页全部对象的调试 JSON
- `pptxcli inspect --slide 0`：复用当前会话输出指定页全部对象的调试 JSON
- `pptxcli show --input ./demo.pptx --slide 0`：渲染指定页预览图
- `pptxcli show --slide 0`：复用当前会话渲染预览图
- `pptxcli show --input ./demo.pptx --slide 0 --annotate`：输出带编号框的标注图
- `pptxcli template create --from ./demo.pptx --name demo_template`：自动启动后台会话并创建模板草稿 JSON
- `pptxcli template create --name demo_template`：复用当前会话创建模板草稿 JSON
- `pptxcli template show --slide 0 --annotate`：复用当前会话输出模板候选标注图
- `pptxcli template add_slide --slide 0 -f "1:title" -f "2:author"`：将当前页字段选择写入当前模板草稿
- `pptxcli template save`：裁剪当前模板对应的原始 PPTX 并生成模板包
- `pptxcli edit create --template demo_template --output ./new.pptx`：创建一个基于模板的编辑草稿
- `pptxcli edit show_template --slide 0`：查看当前编辑模板某一页有哪些 field 需要填充
- `pptxcli edit fill_template --slide 0 -f "1:main title" -f "2:./cover.png"`：向当前编辑草稿追加一页已填充的模板页
- `pptxcli edit fill_template --slide 0 -f "1:main title" --content '{"layout":"vertical","children":[...]}'`：在模板字段之外追加 Container 主体内容
- `pptxcli edit save`：保存当前编辑草稿并生成最终 PPTX

## 会话模式示例

```bash
pptxcli template create --from ./demo.pptx --name quarterly_report
pptxcli template show --slide 0 --annotate --candidates-out ./slide-0.candidates.json
# agent 看标注图后，直接选择字段编号和说明
pptxcli template add_slide --slide 0 \
  -f "1:main title" \
  -f "2:cover image"
pptxcli template save
pptxcli edit create --template quarterly_report --output ./filled.pptx
pptxcli edit show_template --slide 0
pptxcli edit fill_template --slide 0 \
  -f "1:Quarterly Review" \
  -f "2:./cover.png"
pptxcli edit fill_template --slide 0 \
  -f "1:Appendix Title" \
  -f "2:./appendix-cover.png"
pptxcli edit save
```

这组命令会：

- 在第一次模板命令时自动启动单实例后台 server，并预加载当前 PPT
- 在不重复传入 `--input` 的情况下复用当前会话
- 使用 `inspect` 在调试时输出指定页的全部对象
- 使用 LibreOffice headless 将 PPT 转成高还原度预览图
- 使用 `template show --annotate` 提取 text/image 候选对象
- 将 agent 通过命令行选中的字段写入模板草稿 JSON
- 模板页名默认使用 `slide_<index>`
- 裁剪原始 PPTX，仅保留选中的模板页，并生成 `manifest.json`
- 使用 `edit create` 创建编辑中的目标 PPT 草稿，并把 session mode 切换到 `edit_ppt`
- 使用 `edit show_template --slide ...` 查看指定模板页的字段编号、类型和说明
- 使用 `edit fill_template --slide ... --field index:value` 逐页追加已填充的模板页，并校验 text/image 字段
- 使用 `edit fill_template --content <json>` 定义 slide 主体的 Vertical / Horizontal / Grid Container 布局树
- 使用 `edit save` 落盘最终 PPTX，并把 session mode 切回模板提取态
- 后台 server 在空闲 3 分钟后自动退出并清理状态文件

## Container 内容 JSON

`edit fill_template --content` 接收一个 JSON 对象。根节点和子节点都支持嵌套，当前支持三种容器布局：

- `layout: "vertical"`
- `layout: "horizontal"`
- `layout: "grid"`

叶子节点支持：

- `type: "text"`：文本内容
- `type: "image"`：位图文件
- `type: "svg"`：SVG 文件

常用字段包括：

- `bbox`：根容器的绝对区域，单位为 EMU；省略时默认铺满整页
- `padding` / `gap`：容器内边距与子元素间距
- `width` / `height` / `min_width` / `min_height`：尺寸约束
- `grow`：在线性布局中参与剩余空间分配
- `align` / `cross_align`：主轴与交叉轴对齐
- `columns` / `rows` / `column_weights` / `row_weights`：Grid 行列配置
- `style`：文本样式，如 `font_size`、`font_name`、`bold`、`color`

示例：

```bash
pptxcli edit fill_template --slide 0 \
  -f "1:Quarterly Review" \
  --content '{
    "layout": "vertical",
    "bbox": {"x": 731520, "y": 1463040, "w": 5486400, "h": 2743200},
    "gap": 109728,
    "children": [
      {
        "type": "text",
        "height": 548640,
        "text": "Agenda",
        "style": {"font_size": 20, "bold": true}
      },
      {
        "layout": "horizontal",
        "grow": 1,
        "gap": 109728,
        "children": [
          {"type": "text", "grow": 1, "text": "Summary Block"},
          {"type": "image", "width": 1371600, "path": "./cover.png", "fit": "contain"}
        ]
      }
    ]
  }'
```

## 后续任务映射

- 任务 002：补充 PPTX 解包与中间模型
- 任务 003：实现候选检测和视觉标注
- 任务 004：会话化服务与状态管理
- 任务 005：模板确认与 manifest 生成
- 任务 006：模板填充并输出新 PPT
- 任务 007：稳定 Agent 交互协议

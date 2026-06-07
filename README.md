# pptx_cli

`pptx_cli` 是一个面向 Agent 的 PPT 生产命令行工具。首版目标不是让 Agent 直接操作复杂的 PPT 对象模型，而是围绕“模板 + JSON 表单 + 生成/预览”建立稳定工作流。

## 当前状态

当前已完成任务 001，并补上了任务 003 的首版能力：

- 建立了可运行的 CLI 骨架
- 明确了首版技术选型与目录结构
- 提供了一个最小 demo 命令
- 补充了 PPT 解析、写入、预览的技术路线说明
- 支持为指定 slide 生成候选对象 JSON
- 支持通过 LibreOffice 渲染高还原度预览图
- 支持输出目标检测风格的标注框预览图

## 技术选型

- **语言**：Python 3.11+
- **依赖管理**：`uv` + `pyproject.toml` + 本地 `.venv`
- **CLI 框架**：标准库 `argparse`
- **PPT 读写路线**：首选 `python-pptx`
- **数据模型路线**：内部采用 Python `dataclass`；对外 JSON 协议采用显式 schema 文档，后续可升级到 Pydantic
- **预览路线**：通过 LibreOffice headless 将 PPT 转 PDF/图片

更完整的论证见 [docs/tech-selection.md](docs/tech-selection.md)。

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
│       └── show.py
├── tasks/
└── tests/
    ├── test_cli.py
    └── test_show.py
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
- `pptxcli show --input ./demo.pptx --slide 0`：渲染指定页预览图
- `pptxcli show --input ./demo.pptx --slide 0 --annotate`：输出带编号框的标注图

## 任务 003 示例

```bash
pptxcli show --input ./demo.pptx --slide 0 --annotate
```

该命令会：

- 使用 `python-pptx` 提取指定页中的 text/image 候选对象
- 使用 LibreOffice headless 将 PPT 转成高还原度 PDF 预览
- 将对应页渲染为 PNG
- 在 PNG 上叠加类似目标检测任务的醒目标注框，并在框内左上角绘制编号块
- 向标准输出打印候选对象 JSON 和输出图片路径

## 后续任务映射

- 任务 002：补充 PPTX 解包与中间模型
- 任务 003：实现候选检测和视觉标注
- 任务 004：生成模板包与 manifest
- 任务 005：模板填充并输出新 PPT
- 任务 006：局部修改与预览
- 任务 007：稳定 Agent 交互协议

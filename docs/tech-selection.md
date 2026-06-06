# 技术选型说明

## 1. 实现语言与依赖管理

### 结论

首版选择 **Python 3.11+**，使用 **`uv` + `pyproject.toml`** 作为依赖管理方案，并在仓库根目录维护 **`.venv`** 虚拟环境。

### 选择原因

- Python 在 Office / OOXML 生态里工具更成熟，资料也更丰富
- 对 Agent 友好，适合快速输出 JSON、中间模型和调试工具
- 便于后续把命令拆成“解析 / 检查 / 生成 / 预览”多个子命令
- 可以较低成本接入图像标注、文件处理、JSON schema 校验等能力
- `uv` 足够轻量，解析依赖、创建虚拟环境、锁定依赖都在同一套工作流内

### 落地方式

- 使用 `uv venv .venv` 在项目根目录创建虚拟环境
- 使用 `uv sync` 根据 `pyproject.toml` 和 `uv.lock` 同步环境
- 后续新增依赖统一使用 `uv add <package>`
- 开发命令统一使用 `uv run <command>`

### 暂不选择 Go / Node.js 的原因

- Go 的 PPTX 高层生态较弱，后续模板填充和对象操作会更重
- Node.js 虽适合 CLI，但在 PPT OOXML 操作和桌面预览链路上没有明显优势
- MVP 更需要“可调试、可演进”，而不是首日单二进制分发

## 2. CLI 框架

### 结论

首版使用 **标准库 `argparse`**。

### 原因

- 当前任务目标是先搭建 CLI 骨架并验证命令形态，`argparse` 足够支撑
- 零额外依赖，保证仓库开箱即可运行 `./pptxcli --help`
- 子命令结构清晰，后续迁移到 Typer/Click 的成本可控

### 命令结构草案

```text
pptxcli
├── inspect        # 解析/查看 PPT 基础结构
├── template       # 模板相关能力
│   ├── detect     # 候选检测与标注
│   ├── build      # 生成模板 manifest
│   ├── show       # 输出表单 schema / demo
│   ├── fill       # 用表单填充模板
│   └── modify     # 局部替换或插入 slide
├── preview        # 导出预览图
├── demo           # 输出 demo 数据
└── tech           # 输出技术路线摘要
```

## 3. JSON schema / 数据模型方案

### 结论

- **内部模型**：使用 `dataclass` 构建 slide / shape / field 的中间模型
- **外部协议**：使用 JSON 文件作为 Agent 交互介质，并保持 schema 显式、稳定、可版本化

### 设计原则

- 字段名稳定，便于 Agent 重试和 patch
- 模型分层：`raw ppt structure -> candidate model -> template manifest -> fill form`
- 每一步都可导出 JSON，便于调试与人工核验

### 首版核心模型

#### PPT 结构中间模型

```json
{
  "slides": [
    {
      "slide_id": "slide-1",
      "index": 0,
      "shapes": [
        {
          "shape_id": "shape-1",
          "kind": "text",
          "bbox": {"x": 10, "y": 10, "w": 300, "h": 80},
          "text": "Quarterly Review"
        }
      ]
    }
  ]
}
```

#### 模板表单模型

```json
{
  "template_path": "template.pptx",
  "slides": [
    {
      "slide": "title-page",
      "fields": {
        "main_title": "AI Strategy",
        "sub_title": "2026 Q2",
        "author": "Team Solo"
      }
    }
  ]
}
```

### 为什么暂不在任务 001 引入 Pydantic

- 当前主要目标是完成 CLI 初始化和技术路线落地
- 在模型边界还未最终稳定前，先用 `dataclass` + JSON 文档更轻量
- 到任务 002/004 时，如 schema 已稳定，再引入 Pydantic 生成校验和 schema 更合适

## 4. PPT 解析与写入技术路线

### 结论

首选 **`python-pptx`** 作为 MVP 的 PPT 读写基础库。

### 可行性判断

- 能读取演示文稿、slide、shape、文本框、图片等基础对象
- 能修改文本内容，替换图片，复制模板页的方案也有可行实现空间
- 对 MVP 需要的 text/image 替换是够用的

### 风险与边界

- 对复杂对象（图表、SmartArt、动画、母版细节）的支持有限
- slide 深复制、关系引用、媒体替换等细节要谨慎处理
- 某些“模板保持原样”的需求可能需要直接操作底层 OOXML 补强

### 策略

- 任务 002 先建立独立中间模型，避免业务逻辑直接耦合 `python-pptx`
- 对于高层 API 不足的地方，补充 zip + XML 级别读取能力
- MVP 明确只支持 text/image，不提前承诺复杂对象

## 5. 预览技术路线

### 结论

预览采用 **LibreOffice headless** 转换链路。

### 推荐命令

```bash
libreoffice --headless --convert-to pdf input.pptx --outdir ./out
```

后续可再补：

- `pdftoppm`：将 PDF 转 PNG/JPG
- ImageMagick：做裁剪、编号叠加和标注合成

### 原因

- LibreOffice 是最现实的跨平台命令行预览基础设施
- 相比自行渲染 PPT，复杂度和维护成本更低
- 能较好支撑“导出整页预览图供 Agent 复核”的核心场景

### 风险

- 机器需安装 LibreOffice
- 不同环境下字体缺失会影响渲染一致性
- 对极复杂动画和过渡效果不作为 MVP 保障范围

## 6. 初始化阶段结论

任务 001 结束后，项目进入“可运行但功能未实现”的状态：

- CLI 入口已经稳定
- 未来子命令空间已经预留
- 技术路线明确聚焦 `python-pptx + JSON + LibreOffice`
- 项目结构可直接承接任务 002 到任务 007

# 任务 007：Container 布局与嵌套排版

## 目标

引入 Container 布局抽象，支持 Vertical / Horizontal / Grid 三种布局类型，用统一的属性控制子元素的大小、间距、对齐与嵌套关系，简化 PPT 的版式生成能力。

## 主要工作

- 设计 Container 中间模型，支持 `vertical`、`horizontal`、`grid` 三种类型
- 定义 Container 的核心属性，如尺寸约束、padding、gap、alignment、cross alignment、grid 行列配置等
- 支持 Container 子节点为文本、图片、SVG、以及其他 Container
- 设计布局求解规则，将抽象布局树转换为 PPT 中的具体几何位置和尺寸
- 明确 Container 与模板字段、表单 JSON 的结合方式
- 处理嵌套 Container、最小尺寸、溢出、拉伸、等比分配等边界情况
- 为后续局部修改与预览提供可调试的布局中间结果

## 交付物

- Container 数据结构定义
- Container 布局协议或 JSON schema
- 布局求解与几何落位说明
- 示例输入输出，覆盖 Vertical / Horizontal / Grid 与嵌套场景

## 验收标准

- Agent 可通过结构化 JSON 描述 Container 布局树
- 工具可将 Container 布局稳定转换为 PPT 中的元素位置与尺寸
- 支持 Container 中混合文本、图片、SVG 与嵌套 Container
- 常见版式场景可在不直接指定大量绝对坐标的情况下完成排版
- 在编辑 ppt 时，可以通过 pptxcli edit fill_template --slide 0 -f "1:xxxx" -f "2:xxxx" -f "3:xxxx" --content "{}" 来填充内容。也就是说，需要添加 --content 参数，其值是一个 json 串，可以解析成嵌套的 Container 的对象，从而定义 slide 的主体内容。
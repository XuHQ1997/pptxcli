# 任务 005：模板确认与 manifest 生成

## 目标

接收 Agent 的决策 JSON，将候选页面固化为模板定义，生成 `template.pptx + manifest.json`。

## 主要工作

- 定义 decision JSON schema
- 定义 manifest schema
- 将候选对象映射为正式字段
- 支持多页面模板汇总与命名
- 基于当前会话状态读取已解析的 slide 与 candidate 信息

## 交付物

- `template confirm` / `template build` 或等价命令
- manifest 生成逻辑
- 模板包目录约定

## 验收标准

- 能根据决策 JSON 生成可复用模板包
- manifest 中包含字段名、字段类型、shape 绑定关系
- 模板包可被后续填充命令直接使用

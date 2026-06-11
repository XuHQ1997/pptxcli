# Container 格式说明

本文档记录 `pptxcli edit fill_template --content '<json>'` 当前支持的精简版 Container 协议。

这版协议的目标不是“足够灵活”，而是“尽量简单”：

- 最外层节点默认充满模板页的 `content_area`
- 内部元素主要通过 `ratio` 按比例分配空间
- `gap` 也使用比例值，而不是绝对距离
- 不开放 `bbox`、`width`、`height`、`align` 这类绝对控制字段
- 不再兼容 `type=container`、`value`、`content`、`src` 等旧写法
- 默认居中渲染文本和图片

## 入口

```bash
pptxcli edit fill_template --slide 0 \
  -f "1:Quarterly Review" \
  --content '<json>'
```

## 支持的节点

当前只有两类节点：

- 容器节点：通过 `layout` 描述布局
- 叶子节点：通过 `type` 描述内容

### 容器节点

支持三种布局：

- `layout: "vertical"`
- `layout: "horizontal"`
- `layout: "grid"`

### 叶子节点

支持三种内容：

- `type: "text"`
- `type: "image"`
- `type: "svg"`

## 最小结构

一个典型的结构如下：

```json
{
  "layout": "vertical",
  "gap": 0.05,
  "children": [
    {
      "type": "text",
      "ratio": 1,
      "text": "Agenda"
    },
    {
      "layout": "horizontal",
      "ratio": 2,
      "gap": 0.05,
      "children": [
        {
          "type": "text",
          "ratio": 2,
          "text": "Summary Block"
        },
        {
          "type": "image",
          "ratio": 1,
          "path": "./cover.png",
          "fit": "contain"
        }
      ]
    }
  ]
}
```

## 通用规则

### 1. 最外层默认充满 content 区域

调用 `edit fill_template --content ...` 时：

- 工具会读取模板页在 manifest 中记录的 `content_area`
- 最外层节点自动占满这个区域
- 调用方不需要、也不能再手动指定绝对坐标

### 2. 内部空间按 ratio 分配

所有布局下：

- 每个子节点可写 `ratio`
- 未填写时默认 `1`
- `ratio` 可以是单个数字，也可以是 `[横向比例, 纵向比例]`
- 写单个数字时，表示横向和纵向都使用这个值
- 写数组时，在不同布局下按对应轴读取
- `gap` 也使用比例值，表示每个相邻间隔占父容器主轴的比例

具体规则：

- `vertical` 按子节点的纵向比例分配高度
- `horizontal` 按子节点的横向比例分配宽度
- `grid` 会按子节点比例推导列宽和行高
  - 每一列取该列所有子节点横向比例中的最大值
  - 每一行取该行所有子节点纵向比例中的最大值

例如：

- `gap: 0.05` 表示每个间隔占父容器主轴的 5%
- 在 `vertical` 下按容器高度计算
- 在 `horizontal` 下按容器宽度计算
- 在 `grid` 下横向间隔按容器宽度计算，纵向间隔按容器高度计算

例如：

```json
{
  "layout": "horizontal",
  "children": [
    {"type": "text", "ratio": 2, "text": "Left"},
    {"type": "text", "ratio": 1, "text": "Right"}
  ]
}
```

表示左右宽度按 `2:1` 分配。

再例如：

```json
{
  "layout": "grid",
  "columns": 2,
  "children": [
    {"type": "text", "ratio": [2, 1], "text": "A1"},
    {"type": "text", "ratio": [1, 1], "text": "A2"},
    {"type": "text", "ratio": [3, 2], "text": "B1"},
    {"type": "text", "ratio": [4, 5], "text": "B2"}
  ]
}
```

这里会得到：

- 第 1 列比例 `max(2, 3) = 3`
- 第 2 列比例 `max(1, 4) = 4`
- 第 1 行比例 `max(1, 1) = 1`
- 第 2 行比例 `max(2, 5) = 5`

### 3. 默认居中

当前版本不再暴露布局对齐字段。

- 文本默认水平居中、垂直居中
- 图片默认在分配到的区域内居中放置
- 如果图片 `fit=contain`，会等比缩放并居中

## 容器节点字段

### Vertical / Horizontal

| 字段 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| `layout` | `string` | 是 | `vertical` 或 `horizontal` |
| `children` | `array` | 是 | 子节点列表 |
| `gap` | `number` | 否 | 子节点间距比例，如 `0.05` |
| `ratio` | `number \| [number, number]` | 否 | 作为父节点子元素时占比；单值表示双轴同值，数组表示 `[横向, 纵向]` |
| `name` | `string` | 否 | 调试名 |

### Grid

| 字段 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| `layout` | `string` | 是 | 固定为 `grid` |
| `children` | `array` | 是 | 子节点列表，按顺序填充网格 |
| `columns` | `number` | 否 | 列数；不写时自动推导 |
| `gap` | `number` | 否 | 网格间距比例，如 `0.05` |
| `ratio` | `number \| [number, number]` | 否 | 作为父节点子元素时占比；子节点的 `ratio` 同时参与列宽/行高推导 |
| `name` | `string` | 否 | 调试名 |

说明：

- `grid` 当前按顺序自动排布子节点
- 不支持 `row`、`column`、`row_span`、`col_span`

## 叶子节点字段

### Text

| 字段 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| `type` | `string` | 是 | 固定为 `text` |
| `text` | `string` | 是 | 文本内容 |
| `ratio` | `number \| [number, number]` | 否 | 作为父节点子元素时占比；单值表示双轴同值，数组表示 `[横向, 纵向]` |
| `name` | `string` | 否 | 调试名 |
| `style` | `object` | 否 | 文本样式 |

当前支持的 `style` 字段：

- `font_name`
- `font_size`
- `bold`
- `italic`
- `underline`
- `color`

示例：

```json
{
  "type": "text",
  "text": "Agenda",
  "style": {
    "font_name": "Arial",
    "font_size": 20,
    "bold": true,
    "color": "#333333"
  }
}
```

### Image

| 字段 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| `type` | `string` | 是 | 固定为 `image` |
| `path` | `string` | 是 | 图片路径 |
| `ratio` | `number \| [number, number]` | 否 | 作为父节点子元素时占比；单值表示双轴同值，数组表示 `[横向, 纵向]` |
| `fit` | `string` | 否 | `fill` / `contain` / `cover`，默认 `contain` |
| `name` | `string` | 否 | 调试名 |

示例：

```json
{
  "type": "image",
  "path": "./cover.png",
  "ratio": 1,
  "fit": "contain"
}
```

### SVG

`svg` 和 `image` 使用同样的字段：

```json
{
  "type": "svg",
  "path": "./chart.svg",
  "ratio": 1
}
```

说明：

- 当前 SVG 与图片共用一条渲染路径
- `.svg` 当前按目标区域直接放置

## 明确不再支持的字段和写法

为了保持协议简单，以下内容不再支持：

- `bbox`
- `x` / `y` / `w` / `h`
- `width` / `height`
- `min_width` / `min_height`
- `max_width` / `max_height`
- `align` / `cross_align`
- `padding`
- `rows`
- `row` / `column`
- `row_span` / `col_span`
- `type: "container"`
- 文本节点使用 `value` / `content`
- 图片或 SVG 使用 `src` / `value`

如果继续传这些旧字段，当前实现会直接报错。

## 当前校验规则

当前实现会做如下校验：

- `--content` 必须是 JSON 对象
- 根节点必须使用 `layout`
- 容器节点必须包含 `children`
- `layout` 只能是 `vertical` / `horizontal` / `grid`
- `type` 只能是 `text` / `image` / `svg`
- `ratio` 必须是正数，或长度为 `2` 的正数数组
- 文本节点必须提供 `text`
- 图片和 SVG 节点必须提供存在的 `path`
- `style` 中只能出现当前支持的字体样式字段

## 返回结果中的调试信息

执行 `edit fill_template --content ...` 后，返回 JSON 中会包含：

- `content_count`
- `content_layout`
- `rendered_content`

其中：

- `content_layout` 表示布局求解后的树结构
- 每个节点都带最终 `bbox`
- `rendered_content` 记录真正写入 PPT 的叶子节点和 `shape_id`

## 完整示例

```json
{
  "layout": "vertical",
  "name": "page_body",
  "gap": 0.05,
  "children": [
    {
      "type": "text",
      "name": "section_title",
      "ratio": 1,
      "text": "Agenda",
      "style": {
        "font_name": "Arial",
        "font_size": 20,
        "bold": true,
        "color": "#222222"
      }
    },
    {
      "layout": "horizontal",
      "name": "content_row",
      "ratio": 2,
      "gap": 0.05,
      "children": [
        {
          "type": "text",
          "name": "summary_block",
          "ratio": 2,
          "text": "Summary Block",
          "style": {
            "font_size": 14,
            "color": "#444444"
          }
        },
        {
          "type": "image",
          "name": "hero_image",
          "ratio": 1,
          "path": "./cover.png",
          "fit": "contain"
        }
      ]
    }
  ]
}
```

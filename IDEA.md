我想编写一个辅助 AGENT 制作 PPT 的命令行工具，支持预览 PPT、直接编辑 PPT、从已有 PPT 中制作模板、根据模板创建 PPT 等功能。 

我结合用法说一下我设想的实现.
1. 用户对 agent 下达指令，要求根据 existing.pptx 制作一个 new.pptx
2. agent 根据我们的 skill 提示，会利用我们的 cli 工具从 existing.pptx 制作出一个模板。
3. agent 使用这个模板和我们的 cli 工具，进行内容填充，生成 new.pptx 文件。

这里面最复杂的点在于如何辅助 agent 从已有 ppt 中提取出模板。这里的模板文件格式是可以完全由我们定义，可以是一个单独的 pptx 文件，也可以是一个 pptx 文件加上一个 json 文件。

我们的工具支持输入一个模板文件，输出一个 json 数据，这个 json 数据的作用是一个表单，供 agent 填写。比如：
```pptxcli template show template.pptx```
工具解析 template.pptx，并且输出一个 json。json 的样例是：
```json
{
    "template_path": "template.pptx",
    "template_slides": {
        "title page":  {
            "main title": {
                "description" : "main title of the presentation",
                "type": "text"
            },
            "sub title": {
                "description" : "sub title of the presentation",
                "type": "text"
            },
           "author": {
                "description" : "author of the presentation",
                "type": "text"
            },
        },
        "two column content page": {
            "content 1": {
                "description": "content 1",
                "type": "text"
            },
            "image 1": {
                "description": "image 1",
                "type": "image"
            },                
            "content 2": {
                "description": "content 2",
                "type": "text"
            },
            "image 2": {
                "description": "image 2",
                "type": "image"
            }
        }
    }
}
```

agent 填写这个表单就是在 fields 相应的位置补充上 content。比如：
```json
{
    "template_path": "template.pptx",
    "slides": [
        {
            "slide": "title page",
            "fields": {
                "main title": "main title of the presentation",
                "sub title": "sub title of the presentation",
                "author": "author of the presentation"
            }
        },
        {
            "slide": "two column content page",
            "fields": {
                "content 1": "content 1",
                "image 1": "path/to/image 1",
                "content 2": "content 2",
                "image 2": "path/to/image 2"
            }
        }
    ]
}
```

之后，我们的工具根据表单和模板文件，生成最终的 pptx 文件。这样 agent 只需处理 json 表单，简化 agent 操作 ppt 的过程。比如：
```pptxcli template fill --output new.pptx --form json_data.json```

然后 agent 也可以调整某一页的内容：
```pptxcli template modify --slide 0 --form json_data.json```
表示修改第一页的内容，输入的 json 可以如下，表示将 slide 0 替换成两个 slide，一个是 title page，一个是 two column content page。
```json
{
    "template_path": "template.pptx",
    "slides": [
        {
            "slide": "title page",
            "fields": {
                "main title": "main title of the presentation",
                "sub title": "sub title of the presentation",
                "author": "author of the presentation"
            }
        },
        {
            "slide": "two column content page",
            "fields": {
                "content 1": "content 1",
                "image 1": "path/to/image 1",
                "content 2": "content 2",
                "image 2": "path/to/image 2"
            }
        }
    ]
}
```

为了便于 agent 预览，我们再支持预览命令。比如：
```pptxcli preview --input new.pptx --slide 0 --output preview.jpg```


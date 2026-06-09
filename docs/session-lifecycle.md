# 会话生命周期

## 目标

`pptx_cli` 通过“首次模板命令自动启动会话 -> 多轮命令复用 -> 空闲自动退出”的模式复用当前 PPT 的加载与解析现场，避免每次命令都重复传入完整上下文。

## 命令

### 1. 自动启动会话

```bash
pptxcli template create --from ./demo.pptx --name quarterly_report
```

- 自动启动单实例后台 HTTP server
- 由 server 预加载 `Presentation`
- 将监听地址和源文件路径写入状态文件

### 2. 复用会话

在已有会话下，`inspect`、`show`、`template show` 可以省略 `--input`：

```bash
pptxcli inspect --slide 0
pptxcli show --slide 0 --annotate
pptxcli template show --slide 0 --annotate
```

如果对 `inspect/show` 显式传入 `--input`，命令会按一次性本地模式直接执行；如果对 `template show` 显式传入 `--input`，命令会自动启动或复用会话。

### 3. 自动结束会话

- 后台 server 空闲 3 分钟后自动退出
- 删除状态文件
- 丢弃未完成的内部状态与工作进度

## 状态文件

默认状态文件路径：

- 仓库脚本模式：`<pptxcli 所在目录>/.pptxcli-session.json`
- 测试或自定义环境：可通过 `PPTXCLI_STATE_FILE` 覆盖

当前格式约定：

```json
{
  "version": 1,
  "pid": 12345,
  "server_url": "http://127.0.0.1:45123",
  "origin_file": "/abs/path/demo.pptx",
  "created_at": "2026-06-07T12:34:56.000000+00:00"
}
```

字段说明：

- `version`：状态文件版本
- `pid`：后台 server 进程号
- `server_url`：本地监听地址，使用随机端口避免固定端口冲突
- `origin_file`：当前会话加载的 PPT 路径
- `created_at`：server 启动时间

## 异常处理

- 若自动启动时发现已有可用会话，会优先复用同一个源文件对应的会话
- 若状态文件存在但 server 不可达，会先清理失效状态，再允许重新启动
- 若复用命令未找到可用会话，会返回明确错误并提示重新开始模板流程
- 仍保留内部 `finish` 调试命令，用于测试或手动清理会话

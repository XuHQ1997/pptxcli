# 会话生命周期

## 目标

`pptx_cli` 通过 `init -> 多轮命令 -> finish` 的单会话模式复用当前 PPT 的加载与解析现场，避免每次命令都重复传入完整上下文。

## 命令

### 1. 启动会话

```bash
pptxcli init --origin_file ./demo.pptx
```

- 启动单实例后台 HTTP server
- 由 server 预加载 `Presentation`
- 将监听地址和源文件路径写入状态文件

### 2. 复用会话

在已有会话下，`inspect`、`show`、`template detect` 可以省略 `--input`：

```bash
pptxcli inspect --slide 0
pptxcli show --slide 0 --annotate
pptxcli template detect --slide 0 --annotate
```

如果仍显式传入 `--input`，命令会按一次性本地模式直接执行，不依赖当前会话。

### 3. 结束会话

```bash
pptxcli finish
```

- 向后台 server 发送关闭请求
- 删除状态文件
- 清理失效会话的遗留状态

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

- 若 `init` 发现已有可用会话，会拒绝重复启动并提示先执行 `finish`
- 若状态文件存在但 server 不可达，会先清理失效状态，再允许重新 `init`
- 若复用命令未找到可用会话，会返回明确错误并提示重新初始化
- `finish` 遇到失效 server 时会删除陈旧状态文件，避免会话卡死

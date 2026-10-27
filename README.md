# CorelDRAW MCP

通过一个 MCP 服务，让 AI 助手在 Windows 上检查、编辑和保存 CorelDRAW 文档，并完成模板批量处理。

本项目由 **[CorelDRAW Live MCP（CubicDev1/corel-mcp）](https://github.com/CubicDev1/corel-mcp)** 与 **[CorelDRAW Signage MCP（wuguirongsg/coreldraw-mcp）](https://github.com/wuguirongsg/coreldraw-mcp)** 的本地适配版本整合而来。来源、功能继承和新增内容见下方[项目来源](#项目来源与合并方式)。

**推荐与 [coreldraw-layout 排版技能](https://github.com/caoshuo594/coreldraw-layout) 配合使用。** MCP 提供执行工具；技能指导 AI 核对来源、安排版式、保护原图与全文，并完成保存重开验收。安装 MCP 后，再安装技能并按技能描述任务，构成完整的排版工作流。

## 两个仓库如何配合

| 项目 | 负责什么 | 是否可单独使用 |
| --- | --- | --- |
| **本仓库：CorelDRAW MCP** | 连接应用、文档绑定、对象检查、编辑、模板批量、保存与导出 | 可以，调用者需自行安排工作流和验收 |
| **[coreldraw-layout 技能](https://github.com/caoshuo594/coreldraw-layout)** | 任务理解、图文对应、比例与坐标、排版流程、异常处理和交付验收 | 可以，通过 Windows 本地 COM 路径执行 |

搭配后的流程是：**用户描述任务 → AI 读取技能 → 调用 MCP 操作 CorelDRAW → 按技能核查保存文件与预览。**

## 能做什么

本项目整合 Live 的文档会话与编辑能力、Signage 的模板与数据处理能力，共用一个 COM 连接和 STA 执行队列。提供 **18 个直接工具**，另有 **73 项高级功能**按需查询和调用。

- 按完整路径选择文档，避免编辑错误的同名文件。
- 检查对象，用毫米坐标修改文字、字体、颜色和位置。
- 使用修订号检查部分外部变化，将短批次编辑放进一个撤销组。
- 用 CSV / Excel 数据填充命名文本对象，输出多份 CDR。
- 保存可编辑 CDR、导出预览，重新打开磁盘副本检查字体与溢出。
- 在单页工作区中处理框外内容；PNG/JPEG 可按页面、可打印图文或完整工作区范围导出，保留源页面和对象位置。

## 环境与兼容性

| 项目 | 要求 / 已验证环境 |
| --- | --- |
| 系统 | Windows 64 位 |
| Python | 3.11，64 位 |
| CorelDRAW | 已实测 X8，版本 18.0.0.450 |
| 接口 | 本地 COM / pywin32，MCP 使用 stdio |
| 应用生命周期 | 使用前自行启动 CorelDRAW；MCP 不自动启动应用 |

其他 CorelDRAW 版本尚未完成兼容验证，需要确认 ProgID、COM 常量及接口。仓库不包含 CorelDRAW 软件或商业模板。

## 安装 MCP

```powershell
git clone https://github.com/caoshuo594/coreldraw-mcp.git
cd coreldraw-mcp
py -3.11 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

`requirements.txt` 固定了当前实测的直接依赖版本，安装时仍需联网解析传递依赖。本仓库可独立安装，不依赖旧 Live / Signage 的目录或虚拟环境。

先启动 CorelDRAW，再把 [配置示例](mcp-config.example.toml) 中的两条绝对路径改为实际克隆位置，加入 MCP 客户端配置：

```toml
[mcp_servers.coreldraw]
command = 'D:\Tools\coreldraw-mcp\.venv\Scripts\python.exe'
args = ['D:\Tools\coreldraw-mcp\launch.py']
enabled = true
startup_timeout_sec = 60
tool_timeout_sec = 360

[mcp_servers.coreldraw.env]
COREL_ALLOW_RUNNING_DISPATCH = "1"
COREL_ATTACH_POLICY = "existing_only"
COREL_CORELDRAW_APP_NAME = "CorelDRAW.Application.18"
PYTHONUTF8 = "1"
PYTHONUNBUFFERED = "1"
```

上述 ProgID 针对 X8。`COREL_ALLOW_RUNNING_DISPATCH` 是 X8 的连接兼容选项：ROT 查找失败时，仅在检测到一个 CorelDRAW 进程的情况下尝试连接，并检查进程未增加。使用统一服务时停用旧的两个 CorelDRAW MCP。

重新加载客户端后先调用 `corel_status` 和 `list_open_documents`。Codex 的客户端配置见 [官方 MCP 文档](https://learn.chatgpt.com/docs/extend/mcp)。

## 安装配套排版技能

接着按 [coreldraw-layout 安装说明](https://github.com/caoshuo594/coreldraw-layout#安装) 安装技能。以支持该目录的 Codex 环境为例：

```powershell
git clone https://github.com/caoshuo594/coreldraw-layout.git "$env:USERPROFILE\.codex\skills\coreldraw-layout"
```

已有同名技能时先比较并备份，不覆盖本地修改。其他 AI 工具使用各自的技能目录，或直接读取技能的 `SKILL.md`。

安装两者后可向 AI 这样描述任务：

```text
请使用 coreldraw-layout 技能，通过 coreldraw MCP 操作本地 CorelDRAW。
源文件：D:\项目\素材.cdr
参考版式：D:\项目\参考.cdr
输出：D:\项目\交付\展板.cdr
保留指定全文和原图比例，文字与图说保持可编辑。
先核对目标和比例并建立工作副本，保存后重开验收。
```

技能提供工作方法；MCP 工具可用并不代表版式自动达到设计要求。

配套技能默认在同一页面分区、分组组织设计，只有任务明确需要时才分页。页面框与工作区内容范围分别处理；完整工作区预览可使用 `render_preview(scope="workspace", padding_mm=4)`，磁盘重开验收对应使用 `verify_saved_document(preview_scope="workspace")`。具体范围、桌面层行为及矢量输出限制见 [导出范围说明](docs/tools.md#单页工作区与导出范围)。

## 直接调用工具

以下是 MCP 工具调用流程，不是 PowerShell 命令：

1. `list_open_documents` 核对完整路径，再用 `attach_open_document` 绑定。磁盘文件可用 `open_document`。
2. `inspect_document` 读取对象 ID、页面数据和修订号。修改前另存工作副本，并重新检查状态。
3. `apply_operations` 执行短批次编辑，后续操作使用返回的最新修订号。
4. 保存后调用 `verify_saved_document`，再对照源清单检查预览和内容。

```json
{
  "session_id": "从附加结果读取",
  "expected_revision": 0,
  "label": "修改图说",
  "operations": [
    {"op": "text.set_content", "shape_id": "从检查结果读取", "value": "请保持疏散通道畅通"},
    {"op": "text.set_style", "shape_id": "同一对象ID", "font_family": "微软雅黑", "font_size_pt": 18}
  ]
}
```

修订号不能固定为示例中的 `0`。常规编辑使用页面左上角、毫米、Y 向下坐标；另行核对 WorldScale 与真实尺寸。可复制参数见 [修改图说示例](examples/edit-caption.json)。

## 模板批量示例

将模板中的文本对象分别命名为 `room`、`dept`，数据示例见 [batch-data.csv](examples/batch-data.csv)。先调用 `list_feature_tools(name="batch_merge")` 查询参数，再调用 `run_feature_tool`：

```json
{
  "name": "batch_merge",
  "arguments": {
    "template_path": "D:\\Projects\\template.cdr",
    "data_path": "D:\\Projects\\batch-data.csv",
    "output_dir": "D:\\Projects\\output",
    "naming_pattern": "{room}"
  }
}
```

使用模板临时副本，预检文件名并拒绝重复名和已有输出。按对象名称匹配数据列，检查返回的 `matched_fields` 和 `unmatched_fields`，再逐份验收。后续失败会保留已完成文件，重试前先核对输出目录。

## 源码导航

| 文件 | 阅读内容 |
| --- | --- |
| [launch.py](launch.py) | 服务入口；在 COM 线程启动前加载 pandas |
| [server/server.py](server/server.py) | 18 个工具注册、高级功能查询和分发 |
| [connection.py](server/core/connection.py) | 单 STA 线程、COM 连接、重试边界 |
| [sessions.py](server/core/sessions.py) | 文档绑定、修订号、事务和回滚 |
| [process_lock.py](server/core/process_lock.py) | 多个统一服务实例之间的 Windows 互斥锁 |
| [operations.py](server/core/operations.py) | 常规编辑操作与坐标处理 |
| [unified.py](server/tools/unified.py) | 文件打开、模板副本合并、磁盘副本验收 |
| [tests/](tests/) | 连接、坐标、编辑、会话、导出和串行控制测试 |

全部直接工具和高级功能分类见 [工具说明](docs/tools.md)。

## 测试与验证范围

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

当前 28 项自动测试通过。真实 X8 链路验证了中文可编辑文字、字体回读、修订冲突拒绝、失败回滚、保存重开、预览及两条 CSV 批量输出；新增验证覆盖 SVG 导入、指定对象分组、页外对象范围导出，以及含不可打印桌面对象的完整工作区预览。73 项高级功能已验证参数结构，尚未逐项完成真实图纸操作。

- 磁盘验收重新打开字节一致的临时副本，避免关闭用户文档。可选预览为首页，多页视觉验收需补查。
- 部分高级功能使用原生坐标，调用前先查询各自约定。
- 指纹检查不能识别所有字体、颜色或特殊容器变化。
- 统一服务实例间串行 COM 任务，手动操作和其他 COM 控制器不受该锁约束；多个工具调用不是跨聊天原子事务。
- 超时任务可能仍在执行，不能立即重放修改。未知溢出状态不能判为合格。

## 发布下载与问题反馈

源码可直接在线阅读或克隆；历史安装包放在 [Releases](https://github.com/caoshuo594/coreldraw-mcp/releases)，不放在主分支目录中。问题可通过 [Issues](https://github.com/caoshuo594/coreldraw-mcp/issues) 提交，请附版本、复现步骤和错误信息。

## 项目来源与合并方式

| 来源项目 | 原仓库 | 本项目继承的主要能力 |
| --- | --- | --- |
| **CorelDRAW Live MCP** | [CubicDev1/corel-mcp](https://github.com/CubicDev1/corel-mcp) | 指定文档会话、对象检查、修订检查、左上角毫米坐标、短事务编辑、撤销/重做和保存导出 |
| **CorelDRAW Signage MCP** | [wuguirongsg/coreldraw-mcp](https://github.com/wuguirongsg/coreldraw-mcp) | 模板处理、Excel/CSV 数据读取、批量生成、图形/文字/颜色操作和导出检查工具 |

整合前，两套服务分别运行。本项目将本地适配后的执行逻辑放到同一个服务中，共用 COM 连接与 STA 队列；常用编辑保留直接入口，高级功能通过目录查询与统一分发调用。

在来源项目基础上，本项目新增或集中完善了：

- 统一为 18 个直接工具，并提供 73 项高级功能目录和参数查询。
- 多个统一服务实例之间的 Windows 互斥控制与失败修改不自动重放。
- 模板临时副本合并、文件名预检与拒绝覆盖已有输出。
- 磁盘字节副本重开验收，以及对应路径、未知溢出与串行控制测试。
- 本地 X8 连接、保存、字体和单位处理的兼容修补，以及配套技能的使用流程。

整合使用的是带本地兼容修改的工作副本。所记录的来源提交、代码关系与本地修改边界见 [来源记录](docs/provenance.md)。

## 版权与许可

感谢上述项目及原始贡献者。两份来源代码保留的版权声明均署名于深圳市玄熵智能科技有限责任公司；本仓库继续保留原始 [Apache-2.0 许可证](LICENSE) 和 [NOTICE](NOTICE)。本项目属于基于 Windows COM 的第三方 CorelDRAW 集成。

# 项目来源与整合记录

## 两个来源项目

| 项目 | 源码仓库 | 整合时本地工作副本的 Git 基线 |
| --- | --- | --- |
| CorelDRAW Live MCP | [CubicDev1/corel-mcp](https://github.com/CubicDev1/corel-mcp) | [`235affbf`](https://github.com/CubicDev1/corel-mcp/commit/235affbfc76c1fb24bae152f52628d888eb42845) |
| CorelDRAW Signage MCP | [wuguirongsg/coreldraw-mcp](https://github.com/wuguirongsg/coreldraw-mcp) | [`147e093b`](https://github.com/wuguirongsg/coreldraw-mcp/commit/147e093b8537f36a86436e441c4a1e0589fa50eb) |

来源仓库地址由整合时的两个本地 Git remote 核实。上述提交标识用于说明整合基线，实际使用的工作副本包含本地兼容修补；不代表来源项目当前最新版本。

两份本地项目元数据的名称分别是 `coreldraw-live-mcp` 和 `coreldraw-signage-mcp`，版本字段分别为 `0.2.2` 和 `0.1.0`。这些是整合时工作副本的字段，不作为本仓库版本号。

## 代码关系

| 部分 | 来源与整合方式 |
| --- | --- |
| 文档会话、坐标、常规操作及主要绘图工具 | 以本地 Live 工作副本为基础，保留其会话控制与常规编辑逻辑 |
| CSV 数据读取 | 使用本地 Signage 工作副本的标准库 CSV 路径，并保留 Excel 读取能力 |
| 高级功能目录 | 按本地 Signage 注册的 73 项功能整理，对接统一工具模块 |
| 单一服务入口、目录查询与分发 | 本次整合新增，18 个直接工具共用连接与执行队列 |
| 模板隔离合并、磁盘副本验收与进程互斥 | 本次整合新增或重新组织，相关实现在 `tools/unified.py`、`core/process_lock.py` |
| X8 兼容和测试 | 保留已有本地修补，并补充路径隔离、失败不重放、未知溢出与跨进程串行测试 |

两个来源项目有共同的工具模块和版权来源。这里按实际采用的本地代码与整合职责说明贡献，不将每个相同函数重复计为新增能力。

本仓库当前有 21 项自动测试通过，另有 X8 中文编辑、模板批量与保存重开实测。73 项高级功能已检查参数结构，尚未逐项完成真实图纸操作。

## 配套技能

[coreldraw-layout](https://github.com/caoshuo594/coreldraw-layout) 是配套的排版工作方法与验收技能。它通过本 MCP 执行操作，也保留直接 COM 路径；它不是上述两个来源 MCP 之一。

## 原始版权

两份来源工作副本的 `NOTICE` 均保留深圳市玄熵智能科技有限责任公司的版权声明。本仓库保留 [NOTICE](../NOTICE) 与 [LICENSE](../LICENSE)，并在 README 中标明来源和本次改动。

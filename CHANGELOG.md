# Changelog

本文件从 `0.2.6` 起记录面向用户的变更；更早的版本请看 `git log`。

发布时 `publish.yml` 会从本文件对应版本段落生成 GitHub Release notes，
所以**新增版本必须在这里加 `## [<version>] - <date>` 段落**，否则 Release 会退回
用 git 提交信息兜底。

## [Unreleased]

### 变更

- `@deepseek-ai/dsh-tools` 的 peer 范围由 `^0.1.0-rc.5 || ^0.2.0-rc.1` 收窄为
  `^0.2.0-rc.1`，与 `engines.dsh` 的口径一致（`0.1.x` 线从未实测，兼容门实际会拒）。
  运行时行为不变。

## [0.2.6] - 2026-11-21

### 适配 dsh `0.2.2-rc.5`

- **`tools.execute` 改为批量契约**（入参 `{ calls }`、返回 `{ results }`）：集成测试随之改写，断言语义不变。
- 集成测试补 `@deepseek-ai/dsh-tool-executor-native` 装配——rc.5 起 `execute` 之前必须存在 executor，否则抛 `ToolExecutorUnavailableError`。
- devDependencies 由 `0.2.0-rc.2` 升到 `0.2.2-rc.5`（dsh-tools / dsh-system-prompt / cordis `~4.1.0`）。
- 新增 `engines.dsh: ">=0.2.0-rc.1 <0.3.0-0"`，把兼容面钉在 `0.2` 补丁线。

### 升级提示

`0.1.x` 线不再声明支持：兼容门优先读 `engines.dsh`，在 `0.1.x` 宿主上本版会被拒绝。
请先把宿主升到 `0.2` 线（实测 `0.2.2-rc.5`）再安装。

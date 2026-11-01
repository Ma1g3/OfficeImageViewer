# 贡献指南

欢迎提交复现步骤、兼容性结果或代码改进。项目使用 C# 5 / .NET Framework 4.8，构建入口是 `build.ps1`，无需 SDK / NuGet。代码与文档按 [MIT](LICENSE) 许可证提供。

## SDD 流程

1. 修改用户行为前，先更新 [spec.md](specs/001-image-preview/spec.md)。
2. 接口或架构变更更新 [plan.md](specs/001-image-preview/plan.md)。
3. 任务进度更新 [tasks.md](specs/001-image-preview/tasks.md)。
4. 在 [verification.md](specs/001-image-preview/verification.md) 记录验证方法、版本和覆盖边界。

## 本地检查

在项目根目录，用 Windows PowerShell 5.1 构建到独立目录：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\build.ps1 -Test -OutputDirectory work\build
.\work\build\ViewerHarness.exe --smoke .\work\visual
```

所有 `.ps1` 必须保存为 **UTF-8 with BOM**。脚本需要由 Windows PowerShell 5.1 解析，PowerShell 7 检查不能代替：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\ScriptCompatibility.ps1
```

`SetupEngineTests` 使用隔离文件系统和模拟注册检查安装、修复、升级、拒绝及回滚，不更改真实产品注册。注册/安装实测会写入本产品的注册项，只在没有本产品安装的测试环境中执行；已有安装的迁移必须得到用户明确升级授权，不得以测试为由覆盖。普通用户与 AllUsers 范围分别验证。

## 真实 Office 验证

按 [真实宿主验收步骤](specs/001-image-preview/office-validation.md) 使用自建、可公开的合成文档测试。记录 Office 位数、完整构建号和被测 DLL SHA256。

核心测试、模拟 COM 对象、独立查看器或 COM 激活通过，都不等于 Word / PowerPoint 实机通过。没有测试的 Office 位数、构建、DPI 和权限范围请写“待测”。不要把旧版 DLL 的结果当作新版本实测。

## 不变约束

- 不修改真实用户文档，不占用剪贴板，不上传图片。
- COM 入口隔离异常；只有成功显示图片才能取消宿主默认双击。
- 安装和卸载仅操作本产品 GUID / ProgID，不更改 Office / Windows 安全设置。
- 截图、日志和环境快照放在忽略目录 `work/`；分享前检查文档标题、账号、路径和图片内容。

## 提交问题或改动

问题报告请写明复现步骤、预期与实际结果、Windows / Office 版本及位数。使用合成示例替代私人文档。诊断信息可能包含安装路径，公开前请隐去个人信息。

改动说明应包含解决的问题、最终行为、验证结果和未测项。提交前运行 `git status --short`，确认不含 `bin/`、`dist/`、`work/`、`.local/` 或 `validation-results/`。

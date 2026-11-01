# 开发约定
本项目使用 SDD。变更用户行为前先更新 specs/001-image-preview/spec.md，接口/架构更新 plan.md，工作进度更新 tasks.md，验证证据更新 verification.md。
C# 5 / .NET Framework 4.8。build.ps1 是无需 SDK/NuGet 的构建入口。
不得把模拟对象、包解析、截图或成功构建当作 Word/PPT 真实宿主验证。
不修改真实用户文档，不占用剪贴板，不上传图片。COM 入口异常必须隔离；未成功显示图片不得取消宿主双击。
安装脚本仅操作本产品 GUID/ProgID，测试不得覆盖已有安装。未获授权不更改 Office/Windows 安全设置。

PowerShell 脚本必须保存为 UTF-8 with BOM，并用 Windows PowerShell 5.1 解析验证；PowerShell 7 能解析不能替代该检查。

# OfficeImageViewer

双击 Word / PowerPoint 中的图片，在独立窗口里放大查看。适用于 Windows 桌面版 Microsoft Office，中文界面，图片在本地处理。

当前版本（预览） **0.2.1**，增加图形安装、修复和系统卸载入口。采用 C# 5 / .NET Framework 4.8、原生 COM Add-in 和 WPF。

## 普通用户使用

1. 下载发布文件 `OfficeImageViewer-Setup-0.2.1.exe`，双击打开。
2. 安装器自动检测 Word / PowerPoint，默认仅为当前用户安装。根据窗口提示保存并关闭正在运行的 Office，再点击“安装”。
3. 安装成功后重新打开 Word / PPT，双击图片，或点击“图片查看 → 查看图片”。

不需要打开 PowerShell、不需要自己选择 DLL 保存目录；下载的安装 EXE 在安装完成后可移走。程序文件由安装器放到固定程序目录。

修复时再次运行安装 EXE；卸载时打开 **Windows 设置 → 应用 → 已安装的应用 → Office 图片查看器 → 卸载**。Office 内的“图片查看 → 使用帮助与诊断”可查看用法及最近一次预览错误。

已有脚本安装会被识别为升级，保留旧源码和文件。机器级安装继续使用所有用户范围，需要管理员权限；不强制结束 Office。当前安装程序尚未做代码签名，组织策略可能阻止未签名程序运行。

## 功能

- 双击单张图片打开查看器，单击仍正常选图。
- 也可选中图片，点击功能区“图片查看 → 查看图片”。
- 滚轮以指针为中心缩放，按住鼠标拖动平移。
- `Home` 适应窗口，`1` 显示原始像素，`Esc` 关闭。
- 读取文档保存的原图，不修改原文档，不占用剪贴板，不上传图片。

裁剪图片可能显示被裁掉的区域。查看器无法恢复原图已被压缩丢失的细节。

## 环境与兼容性

需要 Windows 桌面版 Word / PowerPoint，以及 .NET Framework 4.8。支持的图片格式为 PNG、JPEG、BMP、GIF、TIFF；GIF 仅显示首帧。

| 环境或场景 | 状态 |
| --- | --- |
| 本机 64 位 Word / PowerPoint | 0.2.1 正常冷启动、所测双击、帮助、系统卸载和重装通过（AllUsers） |
| Word 精确图片范围导出缺图 | 0.1.2 已修复，并在复现文档中真实双击验证 |
| 32 位 Office 真实宿主 | 待测；32 位 COM 激活通过不代表宿主验收 |
| 普通用户 CurrentUser 加载、多屏跨 DPI、其他 Office 构建 | 待测 |
| WPS、Mac、网页版、受保护视图、放映模式 | 不在当前支持范围 |
| 外链图片、组合对象、图表、SVG / EMF | 不在当前支持范围 |

验证环境、版本哈希与覆盖边界见 [验证记录](specs/001-image-preview/verification.md)。普通用户实机验收将在独立电脑进行，未完成前不标为通过。

## 从源码构建

下载或克隆源码，在项目根目录打开 **Windows PowerShell 5.1**：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\build.ps1 -Test -Package
```

构建入口使用 Windows 自带的 .NET Framework 编译器，无需 .NET SDK 或 NuGet。产物为 `bin/Release/OfficeImageViewer.dll` 与 `OfficeImageViewer.Setup.exe`；图形安装程序发布到 `dist/OfficeImageViewer-Setup-0.2.1.exe`。脚本发布包用于开发和故障处理。

如果本目录的插件已经安装并被 Office 加载，开发测试请使用独立输出目录，避免覆盖已安装 DLL：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\build.ps1 -Test -OutputDirectory work\build
```

也可使用 Visual Studio 打开 `OfficeImageViewer.csproj`，需要 .NET Framework 4.8 开发包。

## 开发者脚本安装

以下脚本用于注册型开发部署，不管理图形安装器的文件清单和系统应用条目。图形安装的修复和卸载请使用上面的图形入口。

1. 将发布包解压到固定目录；从源码构建后也可直接在项目根目录安装。
2. 关闭所有 Word 和 PowerPoint 窗口。
3. 用与 Office 相同的 Windows 用户、普通权限 PowerShell 执行：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\install.ps1
```

只有一款 Office 时，添加 `-HostApp Word` 或 `-HostApp PowerPoint`。重新打开 Office，看到“图片查看”选项卡后即可双击图片。安装后的 DLL 目录需要保留。

`Bypass` 仅用于此次脚本进程，不修改系统执行策略。安装脚本只操作本插件的 COM / Office 注册项，不更改 Office 信任中心。当前 DLL 未做 Authenticode 签名；有签名要求的组织需自行签名部署。

### 管理员权限运行的 Office

默认 `CurrentUser` 注册面向普通权限 Office。如果 Office 本身以管理员权限运行，关闭 Office 后，在管理员 PowerShell 显式安装：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\install.ps1 -Scope AllUsers
```

卸载时使用同一范围。安装器发现本产品已注册到其他目录时会拒绝覆盖，请先从原位置卸载。

### 排查与卸载

未出现选项卡或双击无反应时，先运行诊断，检查插件目录和加载情况：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\diagnose.ps1
```

再选中图片尝试“图片查看 → 查看图片”，按钮入口会显示失败原因。

关闭 Word / PowerPoint 后卸载：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\uninstall.ps1
# AllUsers 安装对应使用：
# powershell -NoProfile -ExecutionPolicy Bypass -File .\uninstall.ps1 -Scope AllUsers
```

## 项目结构

| 路径 | 内容 |
| --- | --- |
| `src/Core/` | 图片包解析、资源边界与双击判断 |
| `src/Office/` | COM 接入、Word / PPT 适配与生命周期 |
| `src/Viewer/` | WPF 图片查看器 |
| `setup/` | 图形安装器、注册事务与清单卸载 |
| `tests/` | 核心、模拟宿主、查看器和注册检查 |
| `specs/001-image-preview/` | SDD 规格、方案、任务和验证记录 |
| `build.ps1`、`install.ps1`、`uninstall.ps1` | 构建和安装入口 |
| `bin/`、`dist/`、`work/`、`.local/` | 本地产物，Git 忽略 |

Word 读取所选图片的 Flat OPC 数据；精确范围缺图时，按段落内图片序号和数量核对后回退。PowerPoint 使用 `SaveCopyAs` 创建临时副本，按幻灯片和图片 ID 读取资源，并在 `finally` 中清理。图片限制为 64 MB / 3200 万像素，PPT 临时副本限制为 512 MB。预览失败时保留 Office 默认双击。

## 测试与参与开发

`build.ps1 -Test` 执行 36 项核心测试、14 项模拟宿主测试、22 项安装引擎检查，以及 12 个 PowerShell 脚本的 BOM / Windows PowerShell 5.1 解析检查。安装引擎检查使用隔离文件系统及模拟注册/环境，不能当作系统安装或真实 Office 验证。查看器独立检查：

```powershell
.\bin\Release\ViewerHarness.exe --smoke .\work\visual
```

模拟对象、包解析、截图与成功构建均不能替代 Word / PPT 真实宿主验收。更多说明：

- [贡献指南](CONTRIBUTING.md)：开发流程与验证要求。
- [更新记录](CHANGELOG.md)：版本变更。
- [跨电脑验证指南](验证指南.md)与[验收清单](validation-checklist.md)。
- [行为规格](specs/001-image-preview/spec.md)与[技术方案](specs/001-image-preview/plan.md)。

源码上传 GitHub 时提交源码与文档即可。测试日志、截图、用户文档、Office 安装文件和本地归档不应上传；发布二进制可单独放在 GitHub Releases。

## 许可证

[MIT](LICENSE)。项目为独立实现；WordPictureViewer 仅作为需求参考，未复制其源代码。

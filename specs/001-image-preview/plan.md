# 实现计划

## 公开仓库检查

本机诊断资料与旧产物保留在被 Git 忽略的 `.local/` 中，公开文档只保留匿名验证摘要。脚本兼容检查按项目相对路径排除 `.git`、`.local`、`bin`、`obj`、`dist`、`work` 和 `validation-results`，避免扫描本机归档，也避免项目恰好位于同名父目录时漏掉全部源码脚本。PowerShell BOM 与 Windows PowerShell 5.1 解析要求保持不变。

## 0.2 图形交付

### 0.2.1 默认目录 COM 加载修复

真实 Program Files 安装复现 CLR COM 激活 0x80070002。RegAsm /regfile 只读标准对照显示：CLR 注册 CodeBase 使用程序集自身 CodeBase（空格未转义），不能用 Uri.AbsoluteUri 的 %20 版本。安装器在可卸载的反射专用 AppDomain 中读取目标程序集 CodeBase，避免常驻文件锁；开发安装脚本使用 AssemblyName.GetAssemblyName 读取元数据 CodeBase，不将DLL加载到PowerShell会话。增加空格/中文路径回归与脚本WhatIf后独占读取检查，并以真实32/64位COM激活与Word/PPT冷启动验证，不能仅以字符串断言验收。

增加独立 setup/ 目录，C# 5 / .NET Framework 4.8 WPF 安装器，与插件程序集分开编译。EXE 嵌入插件 DLL 与许可证，安装到 LocalAppData/Programs（CurrentUser）或 Program Files（显式 AllUsers），不依赖运行 PowerShell。构建入口继续为 build.ps1；开发脚本保留。

安装引擎先检测环境、运行中宿主、两种范围和两个注册视图的产品状态；仅在预检通过后写入。文件与本产品注册树先快照，失败恢复；只操作产品 GUID / ProgID 与同 GUID 卸载项。清单记录产品身份、文件哈希和范围；卸载核对身份及固定目录，逐文件删除，保留未知文件。运行中安装器通过临时副本完成维护，避免覆盖自身；临时文件保留清理策略，不用 shell 拼接删除。

现有脚本安装仅在旧 DLL 身份可验证、无跨范围冲突时迁移；旧目录不删除。普通用户路径与管理员路径独立验收，默认安装不更改安全策略。WPF 界面采用现有原生工具风格，状态清晰、主要操作单一；帮助显示用法与本地错误，失败不弹终端。

## 插件实现步骤
1. 固化行为规格与交互选择。
2. 使用 C#、.NET Framework、原生 COM add-in 宿主接入 Word/PPT。选择 COM 而非 VSTO，减少安装依赖；沿用微软对象模型，PowerPoint 普通视图使用下述 UI 线程消息兼容路径，不做跨进程注入或全局鼠标钩子。
3. 核心层处理 OPC/Flat OPC 内嵌图片与资源限制；宿主层负责选图识别及双击/功能区生命周期；WPF 层仅查看图片。
4. 使用 Windows 自带 Framework 编译器构建，避免开发者必须先安装 SDK。保留标准项目文件。
5. 用户级注册 mscoree COM 类及两款 Office Addins；诊断工具验证路径、框架、注册项。
6. 自动化覆盖图片包解析、安全边界、事件取消规则和查看器坐标变换；原生窗口执行交互检查；真实 Office 完成矩阵。

## 公共接口
namespace OfficeImageViewer
- ImagePayload: byte[] Bytes, string Title, string Detail。
- IImageViewer: void ShowImage(ImagePayload image, IntPtr owner)。
- PreviewWindowService: IImageViewer 实现，UI 线程调用，至多一个窗口。
- IHostAdapter: ImagePayload TryGetSelectedImage(); IntPtr OwnerWindow; IDisposable。
- WordHostAdapter(object app), PowerPointHostAdapter(object app)：只读选图提取，不注册事件。
- ImagePackageReader.FromWordOpenXml(string xml, string title)：提取单张选定图片。
- ImagePackageReader.FromPresentation(string path, int slideId, int shapeId, string title)：通过 slide ID、shape ID 与 OPC 关系定位图片；只读取本地已生成的副本。
- PreviewController.TryOpen(Func<ImagePayload>, IImageViewer, IntPtr, out string error)：只有图像存在且成功显示才返回 true。

## 0.1.2 Word 精确范围导出遗漏兼容
真实用户文档复现：选中 InlineShape 的单字符 Range.WordOpenXML 无 drawing/blip/media，所在段落导出则有图片。新增明确的 WordPictureNotFoundException，仅无图片节点触发回退，损坏/外链/大小超限等错误不回退。Word 适配器只读枚举该段落 InlineShapes 的图片/外链图片，通过 Start/End 确定所选图片序号；核心 FromWordParagraphOpenXml 按 wp:inline 中图片顺序和预期数量核对，排除 wp:anchor 浮动图及其他图形。计数或身份不匹配时拒绝。
构建增加可选 OutputDirectory，将修复先编译到独立目录，不覆盖运行中 Office 已加载的DLL。完成测试后仅在文档已保存、Office关闭后更新安装，再恢复原文档和选区；不保存或改写用户文档。


## 跨电脑验证交付
pack-validation.ps1 打包已编译 DLL、查看器、安装卸载、自检、中文指南及验收清单，供目标电脑实测；collect-results.ps1 只采集去除用户名的环境元数据与手工填写清单，不读取文档内容、不自动上传。功能验收仍需真实 Word/PPT 证据。

验收后使用 build.ps1 -PackageOnly 或 pack-validation.ps1 -UseExistingBuild 打包当前已验证字节，避免重新编译产生不同 MVID/哈希。默认命令仍执行构建和测试；打包后逐个比较 DLL 哈希。

## 管理员 Office 的安装范围
真实宿主检查确认本机 Word 为提升权限运行、无关联普通用户令牌；CurrentUser COM 无法加载。安装和卸载脚本增加 Scope=CurrentUser|AllUsers，默认保持 CurrentUser，AllUsers 仅操作本产品 HKLM Classes/Office Addins 项。通过真实 Office 加载和两位数 COM 激活验证该范围，不修改安全设置。

## 真实宿主诊断与窗口归属
PowerPoint Automation 没有可用的 HWND 分派属性；插件应在自身进程内读取实际 PPTFrameClass 窗口，并按活动文档标题/前台根窗口匹配，避免选择另一进程。无可用句柄时查看器居中屏幕。COMAddIn.Object 暴露只读事件计数和阶段诊断，供真实 UI 验收查询，不记录文档内容。

查看器首次显示或更换宿主窗口后，以 MonitorFromWindow/GetMonitorInfo 获取实际目标工作区，以物理窗口边界限制位置与大小，避免按主屏尺寸创建的窗口在较小副屏不可达。同一宿主内复用保留用户位置。边界计算由纯函数回归覆盖；跨显示器 DPI 仍需实机覆盖，不以几何测试代替。

## PowerPoint 普通视图双击兼容
真实 Office 中确认 WindowBeforeDoubleClick 订阅成功但图片双击事件计数仍为零。保留对象模型事件，并增加仅限 PowerPoint 自身 UI 线程的 WH_GETMESSAGE 兼容路径；只读取鼠标按下/抬起消息，在本进程 mdiClass 文档区域、正常编辑视图、单张图片且两次点击选区身份一致时识别双击。仅成功显示图片后吞掉第二次点击；不安装全局鼠标钩子、不注入其他进程、不记录按键或文档内容，断开时卸载。

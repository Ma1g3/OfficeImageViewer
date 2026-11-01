using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Reflection;

[assembly: AssemblyTitle("OfficeImageViewer")]
[assembly: AssemblyDescription("Local image preview for Microsoft Word and PowerPoint")]
[assembly: AssemblyVersion("0.2.1.0")]
[assembly: AssemblyFileVersion("0.2.1.0")]
[assembly: ComVisible(false)]

namespace OfficeImageViewer
{
    [ComVisible(true), Guid("26C7344A-8129-4C5B-A462-034F9B5FC712"),
     ProgId("OfficeImageViewer.AddIn"), ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class AddIn : IDTExtensibility2, IRibbonExtensibility
    {
        private object application;
        private IHostAdapter host;
        private PreviewWindowService viewer;
        private Guid eventInterface;
        private int eventId;
        private BeforeDoubleClick handler;
        private bool busy;
        private bool subscribed;
        private string connectionError;
        private object automationOwner;
        private int doubleClicks, previews, selectionType;
        private string lastError;
        private PowerPointClickBridge clickBridge;
        private int canvasDoubleClicks;

        public string GetStatus()
        {
            PowerPointHostAdapter powerpoint = host as PowerPointHostAdapter;
            return "Connected=" + subscribed + ";DoubleClicks=" + doubleClicks + ";Previews=" + previews +
                ";Owner=" + (host == null ? 0L : host.OwnerWindow.ToInt64()) + ";Selection=" + selectionType + ";Error=" + (lastError ?? connectionError ?? "") +
                (powerpoint == null ? "" : ";Host=" + powerpoint.Diagnostic) +
                (clickBridge == null ? "" : ";Bridge=" + clickBridge.Installed + ";HookError=" + clickBridge.InstallError + ";MouseDowns=" + clickBridge.SurfaceMouseDowns + ";CanvasDoubleClicks=" + canvasDoubleClicks);
        }

        public delegate void BeforeDoubleClick([MarshalAs(UnmanagedType.IDispatch)] object selection, [In, Out, MarshalAs(UnmanagedType.VariantBool)] ref bool cancel);

        public void OnConnection(object app, int connectMode, object addInInstance, ref Array custom)
        {
            try
            {
                Disconnect();
                connectionError = null;
                application = app;
                dynamic office = app;
                string name = Convert.ToString(office.Name);
                if (name.IndexOf("PowerPoint", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    host = new PowerPointHostAdapter(app);
                    eventInterface = new Guid("914934C2-5A91-11CF-8700-00AA0060263B");
                    eventId = 2003;
                }
                else if (name.IndexOf("Word", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    host = new WordHostAdapter(app);
                    eventInterface = new Guid("00020A01-0000-0000-C000-000000000046");
                    eventId = 14;
                }
                else return;
                viewer = new PreviewWindowService();
                handler = HandleDoubleClick;
                ComEventsHelper.Combine(app, eventInterface, eventId, handler);
                subscribed = true;
                PowerPointHostAdapter powerpoint = host as PowerPointHostAdapter;
                if (powerpoint != null) clickBridge = new PowerPointClickBridge(powerpoint.GetSelectedPictureKey, HandleCanvasDoubleClick);
                automationOwner = addInInstance;
                try { if (automationOwner != null) { dynamic automation = automationOwner; automation.Object = this; } } catch { }
            }
            catch
            {
                Disconnect();
                connectionError = "插件未能连接到 Office 的图片事件。请重新启动 Office；如果仍然失败，请重新安装插件。";
            }
        }

        private void HandleDoubleClick(object selection, ref bool cancel)
        {
            ++doubleClicks;
            try { dynamic selected = selection; selectionType = Convert.ToInt32(selected.Type); } catch { selectionType = -1; }
            if (cancel || busy || host == null || viewer == null) return;
            string error;
            if (TryPreview(selection, out error)) cancel = true;
        }

        private bool HandleCanvasDoubleClick()
        {
            if (busy || host == null || viewer == null) return false;
            ++canvasDoubleClicks;
            string error;
            return TryPreview(null, out error);
        }

        private bool TryPreview(object selection, out string error)
        {
            error = null;
            if (host == null || viewer == null) { error = connectionError ?? "插件尚未连接到 Word 或 PowerPoint，请重新启动 Office。"; return false; }
            if (busy) { error = "正在读取图片，请稍候。"; return false; }
            busy = true;
            try
            {
                bool shown = PreviewController.TryOpen(delegate { return selection == null ? host.TryGetSelectedImage() : host.TryGetSelectedImage(selection); }, viewer, host.OwnerWindow, out error);
                lastError = error;
                if (shown) ++previews;
                return shown;
            }
            catch { error = "无法读取当前图片，请确认文档处于普通编辑视图。"; lastError = error; return false; }
            finally { busy = false; }
        }

        public string GetCustomUI(string ribbonID)
        {
            return @"<customUI xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">
              <ribbon><tabs><tab id=""OfficeImageViewerTab"" label=""图片查看"">
              <group id=""ImageViewerGroup"" label=""放大查看"">
              <button id=""ViewSelectedImage"" label=""查看图片"" size=""large"" imageMso=""ZoomToSelection""
                onAction=""OnViewImage"" screentip=""打开所选图片"" supertip=""选择一张图片后打开，支持滚轮缩放、拖动和 Esc 关闭。也可以直接双击图片。""/>
              </group><group id=""ImageViewerSupport"" label=""帮助"">
              <button id=""ImageViewerHelp"" label=""使用帮助与诊断"" imageMso=""Help"" onAction=""OnHelp""/>
              </group></tab></tabs></ribbon></customUI>";
        }

        public void OnViewImage([MarshalAs(UnmanagedType.IDispatch)] object control)
        {
            try
            {
                string error;
                if (!TryPreview(null, out error))
                    MessageBox.Show(error ?? "请选择一张嵌入的图片后重试。", "图片查看", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch { }
        }

        public void OnHelp([MarshalAs(UnmanagedType.IDispatch)] object control)
        {
            try
            {
                string message = "双击一张图片，或选中后点击“查看图片”。\n滚轮缩放，拖动平移，1 原始像素，Home 适应窗口，Esc 关闭。\n\n" +
                    "版本：" + Assembly.GetExecutingAssembly().GetName().Version.ToString(3) +
                    "\n插件连接：" + (subscribed ? "已连接" : "未连接，请重新启动 Office") +
                    "\n最近一次预览：" + (lastError ?? connectionError ?? "无错误记录") +
                    "\n\n无法预览时请在普通编辑视图中单选嵌入的 PNG/JPEG 图片，再点击“查看图片”确认原因。" +
                    "\n安装修复或卸载：Windows 设置 → 应用 → 已安装的应用 → Office 图片查看器。";
                MessageBox.Show(message, "图片查看 · 使用帮助与诊断", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch { }
        }

        public void OnDisconnection(int removeMode, ref Array custom) { Disconnect(); }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { Disconnect(); }

        private void Disconnect()
        {
            if (clickBridge != null) { clickBridge.Dispose(); clickBridge = null; }
            try { if (automationOwner != null) { dynamic automation = automationOwner; automation.Object = null; } } catch { }
            automationOwner = null;
            if (subscribed && application != null && handler != null)
                try { ComEventsHelper.Remove(application, eventInterface, eventId, handler); } catch { }
            subscribed = false;
            if (viewer != null) try { viewer.Dispose(); } catch { }
            if (host != null) try { host.Dispose(); } catch { }
            viewer = null; host = null; handler = null; application = null;
        }
    }
}





using System;
using System.Diagnostics;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Markup;

namespace OfficeImageViewer.Setup
{
    public sealed class SetupWindow : Window
    {
        readonly SetupEngine engine;
        readonly bool uninstall;
        InstallScope scope;
        readonly TextBlock heading, description, status;
        readonly TextBox details;
        readonly Button primary, remove, close;
        readonly CheckBox allUsers;
        bool working, finished;
        public SetupWindow(SetupEngine engine, InstallScope? requestedScope, bool uninstall)
        {
            this.engine = engine; this.uninstall = uninstall;
            Title = "Office 图片查看器 · 安装与维护"; Width = 660; Height = 590; MinWidth = 580; MinHeight = 540;
            WindowStartupLocation = WindowStartupLocation.CenterScreen; FontFamily = new FontFamily("Microsoft YaHei UI"); FontSize = 14;
            Background = Brushes.White; Foreground = new SolidColorBrush(Color.FromRgb(30, 35, 42));
            Grid root = new Grid { Margin = new Thickness(32, 26, 32, 24) };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); Content = root;
            StackPanel intro = new StackPanel(); root.Children.Add(intro);
            heading = new TextBlock { Text = "在 Office 里，双击看大图", FontSize = 24, FontWeight = FontWeights.SemiBold, TextWrapping = TextWrapping.Wrap };
            description = new TextBlock { Text = "为 Word 和 PowerPoint 添加图片查看器。安装完成后重新打开 Office 即可使用。", Margin = new Thickness(0, 12, 0, 24), TextWrapping = TextWrapping.Wrap, LineHeight = 23 };
            intro.Children.Add(heading); intro.Children.Add(description);
            ScrollViewer scroll = new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto }; Grid.SetRow(scroll, 1); root.Children.Add(scroll);
            StackPanel body = new StackPanel(); scroll.Content = body;
            body.Children.Add(new TextBlock { Text = "滚轮缩放  ·  拖动平移  ·  Esc 关闭", Margin = new Thickness(0, 0, 0, 20) });
            allUsers = new CheckBox { Content = "为这台电脑的所有用户安装（需要管理员权限）", Margin = new Thickness(0, 0, 0, 14) };
            body.Children.Add(allUsers);
            status = new TextBlock { TextWrapping = TextWrapping.Wrap, LineHeight = 23, Margin = new Thickness(0, 0, 0, 16) }; body.Children.Add(status);
            details = new TextBox { IsReadOnly = true, TextWrapping = TextWrapping.Wrap, BorderThickness = new Thickness(0), Background = new SolidColorBrush(Color.FromRgb(245, 246, 248)), Padding = new Thickness(12), MinHeight = 180, VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
            System.Windows.Automation.AutomationProperties.SetName(details, "安装环境与使用说明");
            Expander expander = new Expander { Header = "环境检查与使用说明", Content = details, Margin = new Thickness(0, 0, 0, 16) }; body.Children.Add(expander);
            StackPanel buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 20, 0, 0) }; Grid.SetRow(buttons, 2); root.Children.Add(buttons);
            remove = Button("卸载", delegate { new SetupWindow(engine, scope, true).Show(); Close(); }); buttons.Children.Add(remove);
            close = Button("取消", delegate { Close(); }); buttons.Children.Add(close);
            primary = Button(uninstall ? "卸载" : "安装", delegate { Run(); }); primary.IsDefault = true;
            primary.Style = (Style)XamlReader.Parse(@"<Style xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation' TargetType='Button'>
              <Setter Property='Background' Value='#005FB8'/><Setter Property='Foreground' Value='White'/><Setter Property='BorderBrush' Value='Transparent'/><Setter Property='BorderThickness' Value='2'/>
              <Setter Property='Template'><Setter.Value><ControlTemplate TargetType='Button'><Border Background='{TemplateBinding Background}' BorderBrush='{TemplateBinding BorderBrush}' BorderThickness='{TemplateBinding BorderThickness}' CornerRadius='2'><ContentPresenter HorizontalAlignment='Center' VerticalAlignment='Center' RecognizesAccessKey='True' Margin='{TemplateBinding Padding}'/></Border></ControlTemplate></Setter.Value></Setter>
              <Style.Triggers><Trigger Property='IsMouseOver' Value='True'><Setter Property='Background' Value='#004A91'/></Trigger>
              <Trigger Property='IsPressed' Value='True'><Setter Property='Background' Value='#003A73'/></Trigger>
              <Trigger Property='IsKeyboardFocused' Value='True'><Setter Property='BorderBrush' Value='#132E4A'/></Trigger>
              <Trigger Property='IsEnabled' Value='False'><Setter Property='Background' Value='#E1E4E8'/><Setter Property='Foreground' Value='#53585F'/></Trigger></Style.Triggers></Style>"); buttons.Children.Add(primary);
            try
            {
                scope = requestedScope ?? engine.SuggestedScope(); allUsers.IsChecked = scope == InstallScope.AllUsers;
                allUsers.IsEnabled = !uninstall;
                allUsers.Checked += delegate { if (!working) { scope = InstallScope.AllUsers; RefreshSafely(); } };
                allUsers.Unchecked += delegate { if (!working) { scope = InstallScope.CurrentUser; RefreshSafely(); } };
                RefreshSafely();
            }
            catch (Exception error) { primary.IsEnabled = false; remove.Visibility = Visibility.Collapsed; ShowError(error); }
            Closing += delegate(object sender, System.ComponentModel.CancelEventArgs e) { if (working) e.Cancel = true; };
        }
        static Button Button(string text, RoutedEventHandler click)
        { Button button = new Button { Content = text, MinWidth = 100, Height = 38, Padding = new Thickness(16, 0, 16, 0), Margin = new Thickness(8, 0, 0, 0) }; button.Click += click; return button; }
        void Refresh()
        {
            details.Text = engine.Describe(scope);
            bool managed = details.Text.Contains("已安装，可修复"); bool legacy = details.Text.Contains("旧版脚本安装");
            remove.Visibility = !uninstall && managed ? Visibility.Visible : Visibility.Collapsed;
            heading.Text = uninstall ? "卸载 Office 图片查看器" : managed ? "Office 图片查看器已安装" : legacy ? "升级到更方便的安装方式" : "在 Office 里，双击看大图";
            description.Text = uninstall ? "关闭 Word 和 PowerPoint 后继续。卸载将移除插件及其安装文件。目录中的其他文件会保留。" : managed ? "可修复插件注册和程序文件，或从 Windows“已安装的应用”中卸载。" : legacy ? "将插件迁移到固定程序目录，并添加系统卸载入口。原源码目录和旧文件会保留。" : "为 Word 和 PowerPoint 添加图片查看器。安装完成后重新打开 Office 即可使用。";
            primary.Content = uninstall ? "卸载" : managed ? "修复安装" : legacy ? "升级" : "安装";
            status.Foreground = Foreground;
            status.Text = details.Text.Contains("运行中的 Office：无") ? "准备就绪。" : "请先保存文档并关闭 Word / PowerPoint，再继续。";
            if (details.Text.Contains("桌面 Office：未检测到") && !uninstall) { status.Text = "未检测到桌面版 Word 或 PowerPoint，请先安装 Office。"; primary.IsEnabled = false; }
        }
        void RefreshSafely()
        {
            try { primary.IsEnabled = true; Refresh(); }
            catch (Exception error) { primary.IsEnabled = false; remove.Visibility = Visibility.Collapsed; ShowError(error); }
        }
        async void Run()
        {
            if (finished) { Close(); return; }
            if (scope == InstallScope.AllUsers && !SetupEnvironment.Detect().Elevated)
            {
                try { Process.Start(new ProcessStartInfo(Assembly.GetExecutingAssembly().Location, "--scope AllUsers" + (uninstall ? " --uninstall" : "") + " --worker") { UseShellExecute = true, Verb = "runas" }); Close(); }
                catch { status.Text = "管理员授权未完成，未修改安装。可重新点击继续。"; }
                return;
            }
            working = true; primary.IsEnabled = false; remove.IsEnabled = false; close.IsEnabled = false; allUsers.IsEnabled = false;
            status.Foreground = Foreground; status.Text = uninstall ? "正在卸载，请稍候…" : "正在安装，请稍候…";
            try
            {
                await Task.Run(delegate { using (SetupLock gate = new SetupLock()) { if (uninstall) engine.Uninstall(scope); else engine.Install(scope); } });
                finished = true; heading.Text = uninstall ? "已卸载" : "安装完成";
                description.Text = uninstall ? "重新打开 Word / PowerPoint 后，图片查看功能将移除。" : "重新打开 Word 或 PowerPoint，双击一张图片即可放大查看。";
                status.Text = uninstall ? "本产品注册已移除；未知或被改动的文件已保留。" : "滚轮缩放，拖动平移，Esc 关闭。\n以后可在 Windows 设置 → 应用中修复或卸载。";
                details.Text = engine.Describe(scope); allUsers.Visibility = Visibility.Collapsed; remove.Visibility = Visibility.Collapsed; close.Visibility = Visibility.Collapsed; primary.Content = "完成";
            }
            catch (Exception error) { ShowError(error); primary.Content = "重试"; }
            finally { working = false; primary.IsEnabled = true; close.IsEnabled = true; allUsers.IsEnabled = !uninstall && !finished; remove.IsEnabled = true; }
        }
        void ShowError(Exception error)
        { status.Foreground = new SolidColorBrush(Color.FromRgb(160, 35, 35)); status.Text = error.Message + (error is System.IO.IOException ? "\n请确认 Office 已关闭、程序目录可写，然后重试。" : ""); }
    }
}

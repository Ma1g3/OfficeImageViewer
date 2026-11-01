using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using OfficeImageViewer;

public static class ViewerHarness
{
    static int count;
    [STAThread] public static int Main(string[] args)
    {
        Application app=new Application {ShutdownMode=ShutdownMode.OnExplicitShutdown};
        byte[] bytes=CreateImage();
        ImagePayload payload=new ImagePayload(bytes,"示例流程图（测试数据）","原图 · 演示素材，1600 × 1000");
        if(args.Length>0 && args[0]=="--smoke"){
            string output=args.Length>1?args[1]:Path.Combine(Environment.CurrentDirectory,"work");
            Directory.CreateDirectory(output);
            try {
                TestViewport();
                BitmapSource bitmap=PreviewWindowService.DecodeImage(payload);
                Check(bitmap.PixelWidth==1600 && bitmap.PixelHeight==1000,"decoded original dimensions");
                bool failed=false;
                try {PreviewWindowService.DecodeImage(new ImagePayload(new byte[]{1,2,3},"",""));}catch(ImageReadException){failed=true;}
                Check(failed,"corrupt image rejected");
                using(PreviewWindowService service=new PreviewWindowService()){
                    string error;
                    bool opened=PreviewController.TryOpen(delegate{return new ImagePayload(new byte[]{1,2,3},"","");},service,IntPtr.Zero,out error);
                    Check(!opened && error.Contains("图片格式不受支持或内容已损坏"),"viewer error detail reaches ribbon controller");
                }
                PreviewWindow window=new PreviewWindow(); window.SetImage(bitmap,payload.Title,payload.Detail);window.Show();Pump();
                Check(window.IsVisible && window.Viewport.IsFit,"window visible at fit");
                Capture(window,Path.Combine(output,"viewer-large.png"));
                window.SetZoom(1);Check(Math.Abs(window.Viewport.Zoom-1)<0.001,"100 percent");
                window.SetZoom(4);window.PanImage(-100000,-100000);
                Check(window.Viewport.X < 0 && window.Viewport.Y < 0,"image pan reaches bottom right");
                window.FitImage();window.Width=520;window.Height=360;Pump();
                Check(window.Viewport.SourceWidth*window.Viewport.Zoom<=window.Viewport.ViewWidth+1,"small window fit");
                Capture(window,Path.Combine(output,"viewer-small.png"));
                var source=PresentationSource.FromVisual(window);
                window.RaiseEvent(new KeyEventArgs(Keyboard.PrimaryDevice,source,0,Key.Escape){RoutedEvent=Keyboard.PreviewKeyDownEvent});
                Pump(); Check(!window.IsVisible,"Esc closes window");
                using(PreviewWindowService service=new PreviewWindowService()){
                    service.ShowImage(payload,IntPtr.Zero);Pump();
                    int before=app.Windows.Count;
                    service.ShowImage(payload,IntPtr.Zero);Pump();
                    Check(app.Windows.Count==before,"service reuses window");
                }
                Pump();Check(app.Windows.Count==0,"dispose closes window");
                Window firstOwner=new Window {Width=200,Height=100,ShowInTaskbar=false};
                Window secondOwner=new Window {Width=200,Height=100,ShowInTaskbar=false};
                firstOwner.Show();secondOwner.Show();Pump();
                using(PreviewWindowService service=new PreviewWindowService()){
                    service.ShowImage(payload,new System.Windows.Interop.WindowInteropHelper(firstOwner).Handle);Pump();
                    service.ShowImage(payload,new System.Windows.Interop.WindowInteropHelper(secondOwner).Handle);Pump();
                    PreviewWindow preview=app.Windows.OfType<PreviewWindow>().Single();
                    Check(new System.Windows.Interop.WindowInteropHelper(preview).Owner==new System.Windows.Interop.WindowInteropHelper(secondOwner).Handle,"reused preview follows current owner");
                    firstOwner.Close();Pump();Check(preview.IsVisible,"closing old owner preserves new preview");
                }
                secondOwner.Close();Pump();
                for(int i=0;i<20;i++){using(PreviewWindowService service=new PreviewWindowService()){service.ShowImage(payload,IntPtr.Zero);Pump();}}
                Pump();Check(app.Windows.Count==0,"20 open-close cycles");
                Console.WriteLine("Viewer smoke: "+count+" checks passed");app.Shutdown();return 0;
            }catch(Exception ex){Console.WriteLine("FAIL "+ex);app.Shutdown();return 1;}
        }
        PreviewWindow main=new PreviewWindow(); main.SetImage(PreviewWindowService.DecodeImage(payload),payload.Title,payload.Detail);
        main.Closed+=delegate{app.Shutdown();};app.Run(main);return 0;
    }
    static void TestViewport(){
        Rect placed=PreviewPlacement.FitBounds(new Rect(2000,900,1000,740),new Rect(1920,0,800,600));
        Check(placed==new Rect(1920,0,800,600),"preview fits smaller secondary work area");
        placed=PreviewPlacement.FitBounds(new Rect(-1700,-200,500,400),new Rect(-1600,0,1600,900));
        Check(placed==new Rect(-1600,0,500,400),"preview clamps negative monitor coordinates");
        ImageViewport v=new ImageViewport();v.Resize(800,600);v.SetImage(1600,1000);
        Check(Math.Abs(v.Zoom-0.5)<0.001,"fit scale");
        v.ZoomAt(1,400,300);double ix=(400-v.X)/v.Zoom,iy=(300-v.Y)/v.Zoom;
        v.ZoomAt(2,400,300);Check(Math.Abs((400-v.X)/v.Zoom-ix)<0.001 && Math.Abs((300-v.Y)/v.Zoom-iy)<0.001,"cursor anchor");
        v.ZoomAt(999,400,300);Check(v.Zoom==16,"max zoom");
        v.ZoomAt(0.001,400,300);Check(v.Zoom==0.1,"min manual zoom");
        v.SetImage(30000,1000);v.Fit();Check(v.Zoom<0.1,"large image full fit");
        double fitted=v.Zoom,fitX=v.X,fitY=v.Y;
        v.ZoomAt(v.Zoom/1.25,400,300);
        Check(v.IsFit && v.Zoom==fitted && v.X==fitX && v.Y==fitY,"zoom out below manual minimum does not enlarge fit");
        v.ZoomAt(v.Zoom*1.25,400,300);
        Check(!v.IsFit && v.Zoom==0.1,"zoom in from small fit reaches manual minimum");
        v.ZoomAt(1,400,300);v.Pan(-1e6,1e6);Check(v.X==v.ViewWidth-v.SourceWidth && v.Y==0,"pan constrained");
        v.Fit();v.Resize(400,300);Check(v.IsFit && Math.Abs(v.Zoom-400.0/30000)<0.001,"fit follows resize");
    }
    static void Check(bool ok,string name){if(!ok)throw new Exception(name);count++;Console.WriteLine("PASS "+name);}
    static void Pump(){DispatcherFrame frame=new DispatcherFrame();Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.ApplicationIdle,new Action(delegate{frame.Continue=false;}));Dispatcher.PushFrame(frame);}
    static void Capture(Window w,string path){
        FrameworkElement root=(FrameworkElement)w.Content;root.UpdateLayout();
        RenderTargetBitmap image=new RenderTargetBitmap((int)root.ActualWidth,(int)root.ActualHeight,96,96,PixelFormats.Pbgra32);image.Render(root);
        PngBitmapEncoder encoder=new PngBitmapEncoder();encoder.Frames.Add(BitmapFrame.Create(image));using(FileStream f=File.Create(path))encoder.Save(f);
    }
    static byte[] CreateImage(){
        DrawingVisual visual=new DrawingVisual();
        using(DrawingContext c=visual.RenderOpen()){
            c.DrawRectangle(Brushes.White,null,new Rect(0,0,1600,1000));
            c.DrawText(new FormattedText("DOCUMENT IMAGE PREVIEW",System.Globalization.CultureInfo.InvariantCulture,FlowDirection.LeftToRight,new Typeface("Segoe UI"),52,Brushes.DarkSlateGray,1),new Point(90,75));
            c.DrawText(new FormattedText("Synthetic test diagram · 1600 × 1000 pixels",System.Globalization.CultureInfo.InvariantCulture,FlowDirection.LeftToRight,new Typeface("Segoe UI"),24,Brushes.SlateGray,1),new Point(92,150));
            string[] labels={"Select image","Open preview","Zoom & inspect"};
            for(int i=0;i<3;i++){
                Rect box=new Rect(90+i*505,310,410,200);c.DrawRoundedRectangle(new SolidColorBrush(Color.FromRgb(229,238,250)),new Pen(Brushes.SteelBlue,3),box,12,12);
                c.DrawText(new FormattedText(labels[i],System.Globalization.CultureInfo.InvariantCulture,FlowDirection.LeftToRight,new Typeface("Segoe UI"),32,Brushes.DarkSlateBlue,1),new Point(box.X+35,box.Y+80));
                if(i<2)c.DrawLine(new Pen(Brushes.SlateGray,5),new Point(box.Right+15,410),new Point(box.Right+75,410));
            }
            for(int i=0;i<40;i++)c.DrawLine(new Pen(i%5==0?Brushes.SlateGray:Brushes.LightGray,1),new Point(90+i*35,650),new Point(90+i*35,880));
            c.DrawText(new FormattedText("Original pixel detail — Esc returns to your document",System.Globalization.CultureInfo.InvariantCulture,FlowDirection.LeftToRight,new Typeface("Segoe UI"),20,Brushes.SlateGray,1),new Point(90,930));
        }
        RenderTargetBitmap bitmap=new RenderTargetBitmap(1600,1000,96,96,PixelFormats.Pbgra32);bitmap.Render(visual);
        PngBitmapEncoder enc=new PngBitmapEncoder();enc.Frames.Add(BitmapFrame.Create(bitmap));using(MemoryStream m=new MemoryStream()){enc.Save(m);return m.ToArray();}
    }
}





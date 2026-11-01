using System;
using System.Runtime.InteropServices;

namespace OfficeImageViewer
{
    [ComVisible(false)]
    public sealed class ImagePayload
    {
        public ImagePayload(byte[] bytes, string title, string detail)
        {
            if (bytes == null || bytes.Length == 0) throw new ArgumentException("图片内容为空。", "bytes");
            Bytes = bytes;
            Title = title ?? "图片";
            Detail = detail ?? String.Empty;
        }
        public byte[] Bytes { get; private set; }
        public string Title { get; private set; }
        public string Detail { get; private set; }
    }

    [ComVisible(false)]
    public interface IImageViewer
    {
        void ShowImage(ImagePayload image, IntPtr owner);
    }

    [ComVisible(false)]
    public interface IHostAdapter : IDisposable
    {
        ImagePayload TryGetSelectedImage();
        ImagePayload TryGetSelectedImage(object selection);
        IntPtr OwnerWindow { get; }
    }

    [ComVisible(false)]
    public static class PreviewController
    {
        // Return true only when the viewer reports successful display.
        // The host uses that result to cancel Office's default double-click.
        public static bool TryOpen(Func<ImagePayload> getImage, IImageViewer viewer,
            IntPtr owner, out string error)
        {
            error = null;
            if (getImage == null || viewer == null)
            {
                error = "图片查看器尚未就绪。";
                return false;
            }
            try
            {
                ImagePayload image = getImage();
                if (image == null)
                {
                    error = "请先选中一张嵌入式图片，再点击“查看图片”。";
                    return false;
                }
                viewer.ShowImage(image, owner);
                return true;
            }
            catch (ImageReadException ex)
            {
                error = ex.Message;
                return false;
            }
            catch (Exception)
            {
                // Do not expose document paths, Office internals, or image content.
                error = "无法打开这张图片。请确认文件可编辑，且图片格式受支持。";
                return false;
            }
        }
    }

    [ComVisible(false)]
    public class ImageReadException : Exception
    {
        public ImageReadException(string message) : base(message) { }
        public ImageReadException(string message, Exception inner) : base(message, inner) { }
    }

    internal sealed class WordPictureNotFoundException : ImageReadException
    {
        internal WordPictureNotFoundException() : base("未能在所选位置找到嵌入式图片。") { }
    }
}



using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;

namespace OfficeImageViewer
{
    // Every object acquired from an Office property is balanced once. The Application
    // supplied by the host is borrowed; adapters never release it or call Quit().
    internal sealed class OfficeComScope : IDisposable
    {
        private readonly List<object> acquired = new List<object>();

        public object Own(object value)
        {
            if (value != null && Marshal.IsComObject(value)) acquired.Add(value);
            return value;
        }

        public void Dispose()
        {
            for (int i = acquired.Count - 1; i >= 0; --i)
            {
                try { Marshal.ReleaseComObject(acquired[i]); }
                catch (InvalidComObjectException) { }
                catch (ArgumentException) { }
            }
            acquired.Clear();
        }
    }

    public sealed class WordHostAdapter : IHostAdapter
    {
        private object application;
        private const int SelectionInlineShape = 7;
        private const int SelectionFloatingShape = 8;
        private const int InlinePicture = 3;
        private const int InlineLinkedPicture = 4;
        private const int Picture = 13;
        private const int LinkedPicture = 11;

        public WordHostAdapter(object app)
        {
            if (app == null) throw new ArgumentNullException("app");
            application = app;
        }

        public IntPtr OwnerWindow
        {
            get
            {
                if (application == null) return IntPtr.Zero;
                try
                {
                    using (OfficeComScope scope = new OfficeComScope())
                    {
                        dynamic app = application;
                        dynamic window = scope.Own((object)app.ActiveWindow);
                        return new IntPtr(Convert.ToInt32((object)window.Hwnd, CultureInfo.InvariantCulture));
                    }
                }
                catch (Exception) { return IntPtr.Zero; }
            }
        }

        public ImagePayload TryGetSelectedImage() { return TryGetSelectedImage(null); }

        public ImagePayload TryGetSelectedImage(object selectionObject)
        {
            if (application == null) throw new ObjectDisposedException("WordHostAdapter");
            try
            {
                using (OfficeComScope scope = new OfficeComScope())
                {
                    dynamic app = application;
                    dynamic documents = scope.Own((object)app.Documents);
                    if ((int)documents.Count == 0) return null;
                    dynamic selection = selectionObject ?? scope.Own((object)app.Selection);
                    int selectionType = (int)selection.Type;
                    if (selectionType != SelectionInlineShape && selectionType != SelectionFloatingShape)
                        return null;
                    dynamic document = scope.Own((object)selection.Document);
                    string title = Convert.ToString((object)document.Name, CultureInfo.InvariantCulture);

                    if (selectionType == SelectionInlineShape)
                    {
                        dynamic pictures = scope.Own((object)selection.InlineShapes);
                        if ((int)pictures.Count != 1) return null;
                        dynamic picture = scope.Own((object)pictures[1]);
                        int type = (int)picture.Type;
                        if (type == InlineLinkedPicture)
                            throw new ImageReadException("这是一张外链图片。请先将图片嵌入文档，再查看原图。");
                        if (type != InlinePicture) return null;
                        dynamic range = scope.Own((object)picture.Range);
                        string xml = (string)range.WordOpenXML;
                        try { return ImagePackageReader.FromWordOpenXml(xml, title); }
                        catch (WordPictureNotFoundException)
                        {
                            // Some Word documents omit a picture when exporting its
                            // one-character range. A paragraph includes it, but may
                            // include other pictures: match the selected range first.
                            int start = (int)range.Start, end = (int)range.End;
                            dynamic inlineParagraphs = scope.Own((object)range.Paragraphs);
                            dynamic inlineParagraph = scope.Own((object)inlineParagraphs[1]);
                            dynamic expandedRange = scope.Own((object)inlineParagraph.Range);
                            dynamic inlineShapes = scope.Own((object)expandedRange.InlineShapes);
                            int selectedIndex = -1, count = 0, matches = 0;
                            for (int i = 1; i <= (int)inlineShapes.Count; ++i)
                            {
                                dynamic candidate = scope.Own((object)inlineShapes[i]);
                                int candidateType = (int)candidate.Type;
                                if (candidateType != InlinePicture && candidateType != InlineLinkedPicture) continue;
                                dynamic candidateRange = scope.Own((object)candidate.Range);
                                if ((int)candidateRange.Start == start && (int)candidateRange.End == end)
                                { selectedIndex = count; ++matches; }
                                ++count;
                            }
                            if (matches != 1) throw new ImageReadException("无法唯一对应段落中的所选图片，请重新选择一张图片。");
                            return ImagePackageReader.FromWordParagraphOpenXml((string)expandedRange.WordOpenXML,
                                title, selectedIndex, count);
                        }
                    }

                    dynamic shapes = scope.Own((object)selection.ShapeRange);
                    if ((int)shapes.Count != 1) return null;
                    dynamic shape = scope.Own((object)shapes[1]);
                    int shapeType = (int)shape.Type;
                    if (shapeType == LinkedPicture)
                        throw new ImageReadException("这是一张外链图片。请先将图片嵌入文档，再查看原图。");
                    if (shapeType != Picture) return null;
                    int shapeId = (int)shape.ID;
                    string shapeName = (string)shape.Name;
                    dynamic anchor = scope.Own((object)shape.Anchor);
                    dynamic paragraphs = scope.Own((object)anchor.Paragraphs);
                    dynamic paragraph = scope.Own((object)paragraphs[1]);
                    dynamic paragraphRange = scope.Own((object)paragraph.Range);
                    string paragraphXml = (string)paragraphRange.WordOpenXML;
                    // Several floating pictures can share an anchor. The package reader
                    // must match this selected shape; taking the first image is unsafe.
                    return ImagePackageReader.FromWordOpenXml(paragraphXml, title, shapeId, shapeName);
                }
            }
            catch (COMException error)
            {
                throw new ImageReadException("无法从 Word 读取当前图片。请在可编辑文档中选中一张内嵌图片后重试。", error);
            }
        }

        public void Dispose() { application = null; }
    }

    public sealed class PowerPointHostAdapter : IHostAdapter
    {
        private object application;
        private string operation;
        private int selectedType, shapeCount, selectedShapeType, slideCount, errorCode;
        public string Diagnostic { get { return "Stage=" + operation + ",Selection=" + selectedType + ",Shapes=" + shapeCount + ",ShapeType=" + selectedShapeType + ",Slides=" + slideCount + ",HRESULT=" + errorCode.ToString("X8"); } }
        private const int NormalView = 9;
        private const int SlideView = 1;
        private const int ShapeSelection = 2;
        private const int Picture = 13;
        private const int LinkedPicture = 11;
        private const int OpenXmlPresentation = 24;
        private const int MacroEnabledPresentation = 25;
        private const int PreserveFontEmbedding = -2; // msoTriStateMixed: SaveCopyAs default.

        public PowerPointHostAdapter(object app)
        {
            if (app == null) throw new ArgumentNullException("app");
            application = app;
        }

        public IntPtr OwnerWindow
        {
            get
            {
                if (application == null) return IntPtr.Zero;
                try
                {
                    dynamic app = application;
                    using (OfficeComScope scope = new OfficeComScope())
                    {
                        dynamic active = scope.Own((object)app.ActiveWindow);
                        return NativeOfficeWindow.FindPowerPointFrame(Convert.ToString((object)active.Caption, CultureInfo.InvariantCulture));
                    }
                }
                catch (Exception) { return IntPtr.Zero; }
            }
        }

        public ImagePayload TryGetSelectedImage() { return TryGetSelectedImage(null); }

        public ImagePayload TryGetSelectedImage(object selectionObject)
        {
            if (application == null) throw new ObjectDisposedException("PowerPointHostAdapter");
            try
            {
                using (OfficeComScope scope = new OfficeComScope())
                {
                    dynamic app = application;
                    dynamic presentations = scope.Own((object)app.Presentations);
                    if ((int)presentations.Count == 0) return null;
                    dynamic window = scope.Own((object)app.ActiveWindow);
                    int viewType = (int)window.ViewType;
                    if (viewType != NormalView && viewType != SlideView) return null;
                    operation = "Selection"; errorCode = 0;
                    dynamic selection = selectionObject ?? scope.Own((object)window.Selection);
                    selectedType = (int)selection.Type;
                    if ((int)selection.Type != ShapeSelection) return null;
                    // A child of a selected group is deliberately not promoted into a
                    // standalone picture: group handling is outside this version's scope.
                    if (Convert.ToBoolean((object)selection.HasChildShapeRange, CultureInfo.InvariantCulture)) return null;
                    dynamic shapes = scope.Own((object)selection.ShapeRange);
                    shapeCount = (int)shapes.Count;
                    if (shapeCount != 1) return null;
                    dynamic shape = scope.Own((object)shapes[1]);
                    operation = "Shape";
                    int shapeType = (int)shape.Type; selectedShapeType = shapeType;
                    if (shapeType == LinkedPicture)
                        throw new ImageReadException("这是一张外链图片。请先将图片嵌入演示文稿，再查看原图。");
                    if (shapeType != Picture) return null;
                    int shapeId = (int)shape.Id;
                    dynamic slideRange = scope.Own((object)selection.SlideRange);
                    slideCount = (int)slideRange.Count;
                    if (slideCount != 1) return null;
                    dynamic slide = scope.Own((object)slideRange[1]);
                    int slideId = (int)slide.SlideID;
                    dynamic presentation = scope.Own((object)window.Presentation);
                    string title = Convert.ToString((object)presentation.Name, CultureInfo.InvariantCulture);
                    bool hasMacros = Convert.ToBoolean((object)presentation.HasVBProject, CultureInfo.InvariantCulture);
                    string temporaryDirectory = Path.Combine(Path.GetTempPath(), "OfficeImageViewer-" + Guid.NewGuid().ToString("N"));
                    string copyPath = Path.Combine(temporaryDirectory, hasMacros ? "preview.pptm" : "preview.pptx");
                    Directory.CreateDirectory(temporaryDirectory);
                    try
                    {
                        // SaveCopyAs captures unsaved inserted/replaced images. Keep the
                        // current font embedding setting: forcing msoFalse can dirty the
                        // original presentation when it was saved with embedded fonts.
                        // Keep VBA in a macro-enabled copy to avoid conversion prompts;
                        // the copy is only parsed as ZIP and is never opened in Office.
                        operation = "SaveCopyAs";
                        presentation.SaveCopyAs(copyPath, hasMacros ? MacroEnabledPresentation : OpenXmlPresentation, PreserveFontEmbedding);
                        operation = "ReadPackage";
                        return ImagePackageReader.FromPresentation(copyPath, slideId, shapeId, title);
                    }
                    finally
                    {
                        // Only the precise file and private directory we created are removed.
                        // Cleanup failures surface as errors instead of silently retaining a deck.
                        if (File.Exists(copyPath)) File.Delete(copyPath);
                        Directory.Delete(temporaryDirectory, false);
                    }
                }
            }
            catch (COMException error)
            {
                errorCode = error.ErrorCode;
                throw new ImageReadException("无法从 PowerPoint 读取当前图片。请在普通编辑视图选中一张内嵌图片，并确认演示文稿已完整下载。", error);
            }
        }

        public long GetSelectedPictureKey()
        {
            if (application == null) return 0;
            try
            {
                using (OfficeComScope scope = new OfficeComScope())
                {
                    dynamic app = application;
                    dynamic window = scope.Own((object)app.ActiveWindow);
                    int view = (int)window.ViewType;
                    if (view != NormalView && view != SlideView) return 0;
                    dynamic selected = scope.Own((object)window.Selection);
                    if ((int)selected.Type != ShapeSelection || Convert.ToBoolean((object)selected.HasChildShapeRange, CultureInfo.InvariantCulture)) return 0;
                    dynamic shapes = scope.Own((object)selected.ShapeRange);
                    if ((int)shapes.Count != 1) return 0;
                    dynamic picture = scope.Own((object)shapes[1]);
                    if ((int)picture.Type != Picture) return 0;
                    dynamic slides = scope.Own((object)selected.SlideRange);
                    if ((int)slides.Count != 1) return 0;
                    dynamic slide = scope.Own((object)slides[1]);
                    return ((long)(int)slide.SlideID << 32) | (uint)(int)picture.Id;
                }
            }
            catch { return 0; }
        }
        public void Dispose() { application = null; }
    }
}



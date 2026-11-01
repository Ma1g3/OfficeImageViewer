using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Xml;
using System.Xml.Linq;

namespace OfficeImageViewer
{
    /// <summary>Reads only referenced embedded bytes; never extracts archive paths to disk.</summary>
    [ComVisible(false)]
    public static class ImagePackageReader
    {
        public const int MaxImageBytes = 64 * 1024 * 1024;
        public const int MaxXmlCharacters = 8 * 1024 * 1024;
        public const int MaxFlatOpcCharacters = 128 * 1024 * 1024;
        public const long MaxPackageBytes = 512L * 1024 * 1024;
        private const int MaxParts = 10000;
        private const string Pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";
        private const string Rel = "http://schemas.openxmlformats.org/package/2006/relationships";
        private const string Vml = "urn:schemas-microsoft-com:vml";
        private const string Office = "urn:schemas-microsoft-com:office:office";
        private const string Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";

        public static ImagePayload FromWordOpenXml(string xml, string title)
        {
            return FromWordOpenXml(xml, title, null, null);
        }

        public static ImagePayload FromWordOpenXml(string xml, string title, int? shapeId, string shapeName)
        {
            return ReadWordOpenXml(xml, title, shapeId, shapeName, null, 0);
        }

        public static ImagePayload FromWordParagraphOpenXml(string xml, string title, int pictureIndex, int expectedPictureCount)
        {
            return ReadWordOpenXml(xml, title, null, null, pictureIndex, expectedPictureCount);
        }

        private static ImagePayload ReadWordOpenXml(string xml, string title, int? shapeId, string shapeName,
            int? pictureIndex, int expectedPictureCount)
        {
            try
            {
                XDocument package = ReadXml(xml, MaxFlatOpcCharacters);
                if (package.Root == null || package.Root.Name != XName.Get("package", Pkg))
                    throw Failure("Word 未返回有效的图片数据包。");
                Dictionary<string, XElement> parts = new Dictionary<string, XElement>(StringComparer.Ordinal);
                foreach (XElement part in package.Root.Elements(XName.Get("part", Pkg)))
                {
                    if (parts.Count >= MaxParts) throw Failure("文档包含过多资源。");
                    string name = NormalizeEntry((string)part.Attribute(XName.Get("name", Pkg)));
                    if (parts.ContainsKey(name)) throw Failure("文档中存在重复资源路径。");
                    parts.Add(name, part);
                }
                Func<string, XDocument> getXml = delegate(string name)
                {
                    XElement part;
                    if (!parts.TryGetValue(name, out part)) throw Failure("图片所需的文档资源缺失。");
                    XElement data = part.Element(XName.Get("xmlData", Pkg));
                    if (data == null || data.Elements().Count() != 1) throw Failure("图片资源格式无效。");
                    return new XDocument(new XElement(data.Elements().Single()));
                };
                string mainPart = FindOfficeDocument(getXml);
                XDocument document = getXml(mainPart);
                if (document.Root == null || !IsWord(document.Root, "document"))
                    throw Failure("Word 主文档格式不受支持。");
                ResolveAlternateContent(document.Root);
                List<WordPicture> pictures = GetWordPictures(document.Root);
                if (pictureIndex.HasValue)
                {
                    pictures = pictures.Where(p => p.IsInline).ToList();
                    if (expectedPictureCount < 1 || pictures.Count != expectedPictureCount ||
                        pictureIndex.Value < 0 || pictureIndex.Value >= pictures.Count)
                        throw Failure("无法唯一对应段落中的所选图片，请尝试单独选择图片后重试。");
                    pictures = new List<WordPicture> { pictures[pictureIndex.Value] };
                }
                if (shapeId.HasValue || !String.IsNullOrEmpty(shapeName))
                    pictures = pictures.Where(p => p.Matches(shapeId, shapeName)).ToList();
                if (pictures.Count == 0) throw new WordPictureNotFoundException();
                if (pictures.Count != 1) throw Failure("所选位置包含多张图片，请单独选择一张。");
                string imagePart = FindRelatedPart(getXml, mainPart, pictures[0].RelationshipId, "image");
                XElement image;
                if (!parts.TryGetValue(imagePart, out image)) throw Failure("文档中的原图资源缺失。");
                XElement binary = image.Element(XName.Get("binaryData", Pkg));
                if (binary == null) throw Failure("该图片不是可直接读取的内嵌图片。");
                string encoded = binary.Value;
                if (encoded.Length > ((long)MaxImageBytes * 4 / 3 + 4096))
                    throw Failure("图片数据超过 64 MB，无法安全预览。");
                byte[] bytes = Convert.FromBase64String(encoded);
                CheckImageLength(bytes.LongLength);
                return Payload(bytes, title);
            }
            catch (ImageReadException) { throw; }
            catch (Exception ex)
            {
                if (!(ex is XmlException) && !(ex is FormatException) && !(ex is InvalidDataException)
                    && !(ex is ArgumentException)) throw;
                throw new ImageReadException("Word 图片数据损坏或格式不受支持。", ex);
            }
        }

        public static ImagePayload FromPresentation(string path, int slideId, int shapeId, string title)
        {
            try
            {
                using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    if (file.Length > MaxPackageBytes) throw Failure("演示文稿超过 512 MB，无法安全预览。");
                    using (ZipArchive zip = new ZipArchive(file, ZipArchiveMode.Read, false))
                    {
                        if (zip.Entries.Count > MaxParts) throw Failure("文档包含过多资源。");
                        Dictionary<string, ZipArchiveEntry> entries = new Dictionary<string, ZipArchiveEntry>(StringComparer.Ordinal);
                        foreach (ZipArchiveEntry entry in zip.Entries)
                        {
                            if (entry.FullName.EndsWith("/", StringComparison.Ordinal)) continue;
                            string name = NormalizeEntry(entry.FullName);
                            if (entries.ContainsKey(name)) throw Failure("文档中存在重复资源路径。");
                            entries.Add(name, entry);
                        }
                        Func<string, XDocument> getXml = delegate(string name)
                        {
                            ZipArchiveEntry entry;
                            if (!entries.TryGetValue(name, out entry)) throw Failure("图片所需的文档资源缺失。");
                            byte[] data = ReadEntry(entry, MaxXmlCharacters);
                            using (MemoryStream stream = new MemoryStream(data, false))
                            using (XmlReader reader = XmlReader.Create(stream, XmlSettings(MaxXmlCharacters)))
                                return XDocument.Load(reader, LoadOptions.None);
                        };
                        string main = FindOfficeDocument(getXml);
                        XDocument presentation = getXml(main);
                        if (presentation.Root == null || !IsPresentation(presentation.Root, "presentation"))
                            throw Failure("演示文稿格式不受支持。");
                        List<XElement> slides = presentation.Descendants().Where(e =>
                            IsPresentation(e, "sldId") && (string)e.Attribute("id") ==
                            slideId.ToString(CultureInfo.InvariantCulture)).ToList();
                        if (slides.Count != 1) throw Failure("未能定位当前幻灯片。");
                        string slidePart = FindRelatedPart(getXml, main, RelationshipAttribute(slides[0], "id"), "slide");
                        XDocument slide = getXml(slidePart);
                        if (slide.Root == null || !IsPresentation(slide.Root, "sld")) throw Failure("幻灯片数据无效。");
                        ResolveAlternateContent(slide.Root);
                        List<XElement> matches = slide.Descendants().Where(e => IsPresentation(e, "pic") &&
                            e.Descendants().Any(n => IsPresentation(n, "cNvPr") &&
                            (string)n.Attribute("id") == shapeId.ToString(CultureInfo.InvariantCulture))).ToList();
                        if (matches.Count != 1) throw Failure("未能定位所选图片。");
                        List<XElement> blips = matches[0].Descendants().Where(e => IsDrawing(e, "blip")).ToList();
                        if (blips.Count != 1) throw Failure("所选对象不是单张嵌入式图片。");
                        string rid = RelationshipAttribute(blips[0], "embed");
                        if (String.IsNullOrEmpty(rid)) throw Failure("外链图片暂不支持预览，请先将图片嵌入文档。");
                        string imagePart = FindRelatedPart(getXml, slidePart, rid, "image");
                        ZipArchiveEntry image;
                        if (!entries.TryGetValue(imagePart, out image)) throw Failure("文档中的原图资源缺失。");
                        byte[] bytes = ReadEntry(image, MaxImageBytes);
                        CheckImageLength(bytes.LongLength);
                        return Payload(bytes, title);
                    }
                }
            }
            catch (ImageReadException) { throw; }
            catch (Exception ex)
            {
                if (!(ex is XmlException) && !(ex is InvalidDataException) && !(ex is IOException)
                    && !(ex is UnauthorizedAccessException) && !(ex is ArgumentException)) throw;
                throw new ImageReadException("无法读取演示文稿中的原图，文件可能已损坏或不可访问。", ex);
            }
        }

        private static ImagePayload Payload(byte[] bytes, string title)
        {
            return new ImagePayload(bytes, title, "原图 · 包含文档中保留的未裁剪区域");
        }

        private static void CheckImageLength(long size)
        {
            if (size <= 0) throw Failure("文档中的图片内容为空。");
            if (size > MaxImageBytes) throw Failure("图片数据超过 64 MB，无法安全预览。");
        }

        private static byte[] ReadEntry(ZipArchiveEntry entry, int maximum)
        {
            if (entry.Length > maximum) throw Failure("文档资源过大，无法安全预览。");
            using (Stream input = entry.Open())
            using (MemoryStream output = new MemoryStream())
            {
                byte[] buffer = new byte[32768];
                int read;
                while ((read = input.Read(buffer, 0, buffer.Length)) != 0)
                {
                    if (output.Length + read > maximum) throw Failure("文档资源过大，无法安全预览。");
                    output.Write(buffer, 0, read);
                }
                return output.ToArray();
            }
        }

        private static XmlReaderSettings XmlSettings(int maximum)
        {
            return new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null,
                MaxCharactersInDocument = maximum, MaxCharactersFromEntities = 1024 };
        }

        private static XDocument ReadXml(string xml, int maximum)
        {
            if (String.IsNullOrEmpty(xml)) throw Failure("没有可读取的图片数据。");
            if (xml.Length > maximum) throw Failure("文档资源过大，无法安全预览。");
            using (StringReader input = new StringReader(xml))
            using (XmlReader reader = XmlReader.Create(input, XmlSettings(maximum)))
                return XDocument.Load(reader, LoadOptions.None);
        }

        private static string FindOfficeDocument(Func<string, XDocument> getXml)
        {
            XDocument rels = getXml("_rels/.rels");
            List<XElement> matches = Relationships(rels).Where(r => HasType(r, "officeDocument")).ToList();
            if (matches.Count != 1) throw Failure("文档主资源关系无效。");
            return ResolveRelationship("", matches[0]);
        }

        private static string FindRelatedPart(Func<string, XDocument> getXml, string source, string id, string type)
        {
            if (String.IsNullOrEmpty(id)) throw Failure("外链图片暂不支持预览，请先将图片嵌入文档。");
            int slash = source.LastIndexOf('/');
            string folder = slash < 0 ? "" : source.Substring(0, slash + 1);
            string leaf = source.Substring(slash + 1);
            XDocument rels = getXml(folder + "_rels/" + leaf + ".rels");
            List<XElement> matches = Relationships(rels).Where(r => (string)r.Attribute("Id") == id).ToList();
            if (matches.Count != 1 || !HasType(matches[0], type)) throw Failure("文档中的图片引用无效。");
            return ResolveRelationship(source, matches[0]);
        }

        private static IEnumerable<XElement> Relationships(XDocument doc)
        {
            if (doc.Root == null || doc.Root.Name != XName.Get("Relationships", Rel))
                throw Failure("文档资源关系格式无效。");
            return doc.Root.Elements(XName.Get("Relationship", Rel));
        }

        private static bool HasType(XElement relationship, string type)
        {
            string value = (string)relationship.Attribute("Type");
            return value == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/" + type
                || value == "http://purl.oclc.org/ooxml/officeDocument/relationships/" + type;
        }

        private static string ResolveRelationship(string source, XElement relationship)
        {
            string mode = (string)relationship.Attribute("TargetMode");
            if (!String.IsNullOrEmpty(mode) && !String.Equals(mode, "Internal", StringComparison.OrdinalIgnoreCase))
                throw Failure("外链图片暂不支持预览，请先将图片嵌入文档。");
            string target = (string)relationship.Attribute("Target");
            return NormalizeTarget(source, target);
        }

        private static string NormalizeEntry(string name)
        {
            if (String.IsNullOrEmpty(name)) throw Failure("文档资源路径无效。");
            string decoded = DecodePath(name);
            string trimmed = decoded.TrimStart('/');
            if (trimmed.Split('/').Any(p => p == "." || p == ".." || p.Length == 0))
                throw Failure("文档包含不安全的资源路径。");
            return NormalizeTarget("", name);
        }

        private static string DecodePath(string path)
        {
            if (String.IsNullOrEmpty(path)) throw Failure("文档资源路径无效。");
            string decoded = Uri.UnescapeDataString(path);
            if (decoded.IndexOfAny(new char[] { '\\', ':', '?', '#', '\0' }) >= 0 ||
                decoded.Any(c => Char.IsControl(c)) || decoded.StartsWith("//", StringComparison.Ordinal))
                throw Failure("文档包含不安全的资源路径。");
            return decoded;
        }

        private static string NormalizeTarget(string source, string target)
        {
            string decoded = DecodePath(target);
            List<string> segments = new List<string>();
            if (!decoded.StartsWith("/", StringComparison.Ordinal) && !String.IsNullOrEmpty(source))
            {
                int slash = source.LastIndexOf('/');
                if (slash >= 0) segments.AddRange(source.Substring(0, slash).Split('/'));
            }
            foreach (string part in decoded.Split('/'))
            {
                if (part == "" || part == ".") continue;
                if (part == "..")
                {
                    if (segments.Count == 0) throw Failure("文档图片引用超出了数据包范围。");
                    segments.RemoveAt(segments.Count - 1);
                }
                else segments.Add(part);
            }
            if (segments.Count == 0 || decoded.EndsWith("/", StringComparison.Ordinal))
                throw Failure("文档资源路径无效。");
            return String.Join("/", segments);
        }

        private static bool InFamily(XElement element, string localName, string standard, string strict)
        {
            return element.Name.LocalName == localName &&
                (element.Name.NamespaceName == standard || element.Name.NamespaceName == strict);
        }
        private static bool IsWord(XElement e, string name)
        {
            return InFamily(e, name, "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                "http://purl.oclc.org/ooxml/wordprocessingml/main");
        }
        private static bool IsPresentation(XElement e, string name)
        {
            return InFamily(e, name, "http://schemas.openxmlformats.org/presentationml/2006/main",
                "http://purl.oclc.org/ooxml/presentationml/main");
        }
        private static bool IsDrawing(XElement e, string name)
        {
            return InFamily(e, name, "http://schemas.openxmlformats.org/drawingml/2006/main",
                "http://purl.oclc.org/ooxml/drawingml/main");
        }
        private static bool IsDocPr(XElement e)
        {
            return InFamily(e, "docPr", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
                "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing");
        }
        private static string RelationshipAttribute(XElement element, string name)
        {
            return (string)element.Attribute(XName.Get(name, "http://schemas.openxmlformats.org/officeDocument/2006/relationships"))
                ?? (string)element.Attribute(XName.Get(name, "http://purl.oclc.org/ooxml/officeDocument/relationships"));
        }
        private static void ResolveAlternateContent(XElement root)
        {
            foreach (XElement alternate in root.Descendants(XName.Get("AlternateContent", Mc)).Reverse().ToList())
            {
                XElement selected = alternate.Elements(XName.Get("Choice", Mc)).FirstOrDefault()
                    ?? alternate.Element(XName.Get("Fallback", Mc));
                if (selected == null) alternate.Remove();
                else alternate.ReplaceWith(selected.Elements().ToList());
            }
        }
        private sealed class WordPicture
        {
            public string RelationshipId;
            public XElement Marker;
            public bool IsInline;
            public bool Matches(int? shapeId, string name)
            {
                if (Marker == null) return false;
                bool byName = !String.IsNullOrEmpty(name) &&
                    (String.Equals((string)Marker.Attribute("name"), name, StringComparison.Ordinal)
                    || String.Equals((string)Marker.Attribute("id"), name, StringComparison.Ordinal));
                if (!shapeId.HasValue) return byName;
                string value = shapeId.Value.ToString(CultureInfo.InvariantCulture);
                string id = (string)Marker.Attribute("id");
                string spid = (string)Marker.Attribute(XName.Get("spid", Office));
                return byName || id == value || spid == value
                    || id == "_x0000_s" + value || id == "_x0000_i" + value
                    || spid == "_x0000_s" + value || spid == "_x0000_i" + value;
            }
        }
        private static List<WordPicture> GetWordPictures(XElement root)
        {
            List<WordPicture> pictures = new List<WordPicture>();
            foreach (XElement drawing in root.Descendants().Where(e => IsWord(e, "drawing")))
            {
                XElement marker = drawing.Descendants().FirstOrDefault(IsDocPr);
                bool inline = drawing.Elements().Any(e => InFamily(e, "inline",
                    "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
                    "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing"));
                foreach (XElement blip in drawing.Descendants().Where(e => IsDrawing(e, "blip")))
                    pictures.Add(new WordPicture { Marker = marker, IsInline = inline,
                        RelationshipId = RelationshipAttribute(blip, "embed") });
            }
            foreach (XElement pict in root.Descendants().Where(e => IsWord(e, "pict")))
            {
                foreach (XElement shape in pict.Descendants(XName.Get("shape", Vml)))
                {
                    foreach (XElement image in shape.Elements(XName.Get("imagedata", Vml)))
                        pictures.Add(new WordPicture { Marker = shape,
                            RelationshipId = RelationshipAttribute(image, "id") });
                }
            }
            return pictures;
        }
        private static ImageReadException Failure(string message)
        {
            return new ImageReadException(message);
        }
    }
}


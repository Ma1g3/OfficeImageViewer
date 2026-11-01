using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Linq;
using OfficeImageViewer;

public static class CoreTests
{
    const string R = "http://schemas.openxmlformats.org/package/2006/relationships";
    const string O = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const string D = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const string P = "http://schemas.openxmlformats.org/presentationml/2006/main";
    const string WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    static int passed, failed;
    static readonly byte[] A = new byte[] { 1, 2, 3, 4 };
    static readonly byte[] B = new byte[] { 9, 8, 7 };
    public static int Main()
    {
        Test("R01 Word exact referenced bytes", delegate { Equal(ImagePackageReader.FromWordOpenXml(Word(false), "doc").Bytes, A); });
        Test("R01 Word floating exact shape", delegate { Equal(ImagePackageReader.FromWordOpenXml(Word(true), "doc", 9, "Picture 9").Bytes, B); });
        Test("R01 ambiguous Word selection fails", delegate { Reject(delegate { ImagePackageReader.FromWordOpenXml(Word(true), "doc"); }); });
        Test("R01 nonexistent floating shape fails", delegate { Reject(delegate { ImagePackageReader.FromWordOpenXml(Word(true), "doc", 77, "Missing"); }); });
        Test("R01 empty exact Word range identifies paragraph fallback", delegate {
            try { ImagePackageReader.FromWordOpenXml(WordWithoutPictures(),"doc"); }
            catch(ImageReadException error) { Assert(error.GetType().Name=="WordPictureNotFoundException","dedicated missing-picture signal"); return; }
            throw new Exception("expected exact-range missing picture signal");
        });
        Test("R01 Word paragraph recovers selected inline picture", delegate {
            Equal(ImagePackageReader.FromWordParagraphOpenXml(Word(false).Replace("wp:anchor","wp:inline"),"doc",0,1).Bytes,A);
        });
        Test("R01 Word paragraph matches second picture instead of first", delegate {
            Equal(ImagePackageReader.FromWordParagraphOpenXml(Word(true).Replace("wp:anchor","wp:inline"),"doc",1,2).Bytes,B);
        });
        Test("R01 Word paragraph excludes floating pictures", delegate {
            string xml=Word(true).Replace(Drawing(4,"imgA"),Drawing(4,"imgA").Replace("wp:anchor","wp:inline"));
            Equal(ImagePackageReader.FromWordParagraphOpenXml(xml,"doc",0,1).Bytes,A);
        });
        Test("R01 Word paragraph rejects count mismatch", delegate {
            Reject(delegate{ImagePackageReader.FromWordParagraphOpenXml(Word(true).Replace("wp:anchor","wp:inline"),"doc",0,1);});
        });
        Test("R01 Word paragraph rejects invalid index", delegate {
            Reject(delegate{ImagePackageReader.FromWordParagraphOpenXml(Word(false).Replace("wp:anchor","wp:inline"),"doc",1,1);});
        });
        Test("R07 Word external relationship rejected", delegate { Reject(delegate { ImagePackageReader.FromWordOpenXml(Word(false).Replace("Target='media/a.png'", "Target='https://example.invalid/a.png' TargetMode='External'"), "doc"); }); });
        Test("R07 Word path escape rejected", delegate { Reject(delegate { ImagePackageReader.FromWordOpenXml(Word(false).Replace("Target='media/a.png'", "Target='../../a.png'"), "doc"); }); });
        Test("R07 Word DTD rejected", delegate { Reject(delegate { ImagePackageReader.FromWordOpenXml("<!DOCTYPE x [<!ENTITY a SYSTEM 'file:///C:/secret'>]>" + Word(false), "doc"); }); });
        Test("R06 corrupt base64 rejected", delegate { Reject(delegate { ImagePackageReader.FromWordOpenXml(Word(false).Replace("AQIDBA==", "%%invalid%%"), "doc"); }); });
        Test("R01 strict Word namespace supported", delegate {
            string xml=Word(false).Replace(W,"http://purl.oclc.org/ooxml/wordprocessingml/main").Replace(D,"http://purl.oclc.org/ooxml/drawingml/main").Replace(WP,"http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing").Replace(O,"http://purl.oclc.org/ooxml/officeDocument/relationships");
            Equal(ImagePackageReader.FromWordOpenXml(xml,"doc").Bytes,A);
        });
        Test("R02 PPT resolves IDs not part order", delegate { WithDeck("normal", delegate(string p) { Equal(ImagePackageReader.FromPresentation(p, 301, 9, "deck").Bytes, B); }); });
        Test("R02 nonexistent slide fails", delegate { WithDeck("normal", delegate(string p) { Reject(delegate { ImagePackageReader.FromPresentation(p, 302, 9, "deck"); }); }); });
        Test("R02 nonexistent shape fails", delegate { WithDeck("normal", delegate(string p) { Reject(delegate { ImagePackageReader.FromPresentation(p, 301, 88, "deck"); }); }); });
        foreach (string mode in new string[] { "external", "escape", "encodedEscape", "missing", "duplicate", "wrongType", "badZipPath" })
        {
            string scenario=mode;
            Test("R07 PPT rejects "+scenario, delegate { WithDeck(scenario, delegate(string p) { Reject(delegate { ImagePackageReader.FromPresentation(p, 301, 9, "deck"); }); }); });
        }
        Test("R05 PPT package read does not mutate bytes", delegate { WithDeck("normal", delegate(string p) { byte[] before=File.ReadAllBytes(p); ImagePackageReader.FromPresentation(p,301,9,"deck"); Equal(File.ReadAllBytes(p),before); }); });
        Test("R06 malformed ZIP rejected", delegate {
            string p=Path.GetTempFileName();try {File.WriteAllText(p,"not zip");Reject(delegate {ImagePackageReader.FromPresentation(p,301,9,"deck");});}finally {File.Delete(p);}
        });
        Test("R06 success allows cancellation", delegate { FakeViewer v=new FakeViewer(); string error; Assert(PreviewController.TryOpen(delegate{return new ImagePayload(A,"","");},v,IntPtr.Zero,out error),"success"); Assert(v.Calls==1 && error==null,"shown once"); });
        Test("R06 no image keeps default action", delegate { FakeViewer v=new FakeViewer();string error; Assert(!PreviewController.TryOpen(delegate{return null;},v,IntPtr.Zero,out error),"no cancel");Assert(v.Calls==0,"no viewer"); });
        Test("R06 extraction failure keeps default", delegate { FakeViewer v=new FakeViewer();string error; Assert(!PreviewController.TryOpen(delegate{throw new ImageReadException("expected");},v,IntPtr.Zero,out error),"no cancel");Assert(error=="expected","message"); });
        Test("R06 display failure keeps default", delegate { FakeViewer v=new FakeViewer{Fail=true};string error; Assert(!PreviewController.TryOpen(delegate{return new ImagePayload(A,"","");},v,IntPtr.Zero,out error),"no cancel"); });
        Test("R02 native double click accepts same window and position",delegate{var t=new ClickSequence();Assert(!t.Register(new IntPtr(1),100,100,1000,500,4,4),"first");Assert(t.Register(new IntPtr(1),102,98,1300,500,4,4),"second");Assert(!t.Pending,"reset");});
        Test("R02 native double click rejects another window",delegate{var t=new ClickSequence();t.Register(new IntPtr(1),100,100,1000,500,4,4);Assert(!t.Register(new IntPtr(2),100,100,1200,500,4,4),"another window");});
        Test("R02 native double click rejects slow or distant clicks",delegate{var t=new ClickSequence();t.Register(new IntPtr(1),100,100,1000,500,4,4);Assert(!t.Register(new IntPtr(1),100,100,1600,500,4,4),"slow");Assert(!t.Register(new IntPtr(1),110,100,1700,500,4,4),"distant");});
        Test("R02 native double click handles tick wrap",delegate{var t=new ClickSequence();t.Register(new IntPtr(1),0,0,UInt32.MaxValue-100,500,4,4);Assert(t.Register(new IntPtr(1),0,0,100,500,4,4),"wrap");});
        Test("R02 native click reset prevents false double",delegate{var t=new ClickSequence();t.Register(new IntPtr(1),0,0,1000,500,4,4);t.Reset();Assert(!t.Register(new IntPtr(1),0,0,1100,500,4,4),"reset");});
        Console.WriteLine("Core tests: "+passed+" passed, "+failed+" failed");
        return failed==0?0:1;
    }
    static void Test(string name, Action action) {try{action();passed++;Console.WriteLine("PASS "+name);}catch(Exception ex){failed++;Console.WriteLine("FAIL "+name+": "+ex.Message);}}
    public static void Assert(bool condition,string message){if(!condition)throw new Exception(message);}
    static void Equal(byte[] actual,byte[] expected){Assert(actual.SequenceEqual(expected),"byte mismatch");}
    static void Reject(Action action){try{action();}catch(ImageReadException){return;}throw new Exception("Expected ImageReadException");}
    sealed class FakeViewer:IImageViewer {public int Calls;public bool Fail;public void ShowImage(ImagePayload p,IntPtr owner){Calls++;if(Fail)throw new InvalidOperationException();}}
    static string Rel(string id,string type,string target,string extra=""){return "<Relationship Id='"+id+"' Type='"+O+"/"+type+"' Target='"+target+"' "+extra+"/>";}
    static string Rels(string body){return "<Relationships xmlns='"+R+"'>"+body+"</Relationships>";}
    static string XmlPart(string name,string xml){return "<pkg:part pkg:name='/"+name+"'><pkg:xmlData>"+xml+"</pkg:xmlData></pkg:part>";}
    static string BinPart(string name,byte[] bytes){return "<pkg:part pkg:name='/"+name+"'><pkg:binaryData>"+Convert.ToBase64String(bytes)+"</pkg:binaryData></pkg:part>";}
    static string Drawing(int id,string rid){return "<w:drawing><wp:anchor><wp:docPr id='"+id+"' name='Picture "+id+"'/><a:graphic><a:blip r:embed='"+rid+"'/></a:graphic></wp:anchor></w:drawing>";}
    public static string Word(bool multi)
    {
        string body="<w:document xmlns:w='"+W+"' xmlns:wp='"+WP+"' xmlns:a='"+D+"' xmlns:r='"+O+"'><w:body><w:p>"+Drawing(4,"imgA")+(multi?Drawing(9,"imgB"):"")+"</w:p></w:body></w:document>";
        return "<pkg:package xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage'>"+
          XmlPart("_rels/.rels",Rels(Rel("main","officeDocument","word/document.xml")))+
          XmlPart("word/document.xml",body)+XmlPart("word/_rels/document.xml.rels",Rels(Rel("imgA","image","media/a.png")+Rel("imgB","image","media/b.png")))+
          BinPart("word/media/b.png",B)+BinPart("word/media/a.png",A)+"</pkg:package>";
    }
    public static string WordWithoutPictures(){return Word(false).Replace(Drawing(4,"imgA"),"");}
    static void Entry(ZipArchive zip,string name,string value){using(StreamWriter w=new StreamWriter(zip.CreateEntry(name).Open(),new UTF8Encoding(false)))w.Write(value);}
    static void Binary(ZipArchive zip,string name,byte[] bytes){using(Stream w=zip.CreateEntry(name).Open())w.Write(bytes,0,bytes.Length);}
    public static void WithDeck(string mode,Action<string> action)
    {
        string path=Path.Combine(Path.GetTempPath(),"oiv-test-"+Guid.NewGuid().ToString("N")+".pptx");
        try {
            using(FileStream file=File.Create(path))using(ZipArchive z=new ZipArchive(file,ZipArchiveMode.Create)){
                Entry(z,"_rels/.rels",Rels(Rel("office","officeDocument","ppt/presentation.xml")));
                Entry(z,"ppt/presentation.xml","<p:presentation xmlns:p='"+P+"' xmlns:r='"+O+"'><p:sldIdLst><p:sldId id='301' r:id='slideSeven'/></p:sldIdLst></p:presentation>");
                Entry(z,"ppt/_rels/presentation.xml.rels",Rels(Rel("slideSeven","slide","slides/slide7.xml")));
                Entry(z,"ppt/slides/slide7.xml","<p:sld xmlns:p='"+P+"' xmlns:a='"+D+"' xmlns:r='"+O+"'><p:cSld><p:spTree><p:pic><p:nvPicPr><p:cNvPr id='4'/></p:nvPicPr><p:blipFill><a:blip r:embed='imgA'/></p:blipFill></p:pic><p:pic><p:nvPicPr><p:cNvPr id='9'/></p:nvPicPr><p:blipFill><a:blip r:embed='imgB'/></p:blipFill></p:pic></p:spTree></p:cSld></p:sld>");
                string target=mode=="escape"?"../../../outside.png":mode=="encodedEscape"?"%2E%2E/%2e%2e/%2e%2e/outside.png":mode=="external"?"https://example.invalid/b.png":"../media/b.png";
                Entry(z,"ppt/slides/_rels/slide7.xml.rels",Rels(Rel("imgA","image","../media/a.png")+Rel("imgB",mode=="wrongType"?"slide":"image",target,mode=="external"?"TargetMode='External'":"")));
                Binary(z,"ppt/media/a.png",A);
                if(mode!="missing")Binary(z,"ppt/media/b.png",B);
                if(mode=="duplicate")Binary(z,"ppt/media/b.png",A);
                if(mode=="badZipPath")Binary(z,"../../outside.png",A);
            }
            action(path);
        }finally {if(File.Exists(path))File.Delete(path);}
    }
}



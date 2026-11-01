using System;
using System.IO;
using OfficeImageViewer;

public static class HostAdapterTests
{
    static int passed,failed;
    public static int Main(){
        Test("Word event selection overrides stale app selection",delegate{
            FakeWord app=Word(0);using(WordHostAdapter a=new WordHostAdapter(app)){
                var selected=Word(7).Selection;
                CoreTests.Assert(a.TryGetSelectedImage(selected).Bytes[0]==1,"callback selected image");
                CoreTests.Assert(a.TryGetSelectedImage()==null,"ribbon uses current selection");
            }
        });
        Test("Word text passes through",delegate{using(var a=new WordHostAdapter(Word(0))){CoreTests.Assert(a.TryGetSelectedImage()==null,"no image");}});
        Test("Word multiple inline shapes pass through",delegate{var app=Word(7);app.Selection.InlineShapes=new FakeCollection(new FakePicture(),new FakePicture());using(var a=new WordHostAdapter(app)){CoreTests.Assert(a.TryGetSelectedImage()==null,"multi");}});
        Test("Word non-picture inline passes through",delegate{var app=Word(7);((FakePicture)app.Selection.InlineShapes[1]).Type=1;using(var a=new WordHostAdapter(app)){CoreTests.Assert(a.TryGetSelectedImage()==null,"chart");}});
        Test("Word floating shape matches exact image",delegate{var app=Word(8);using(var a=new WordHostAdapter(app)){CoreTests.Assert(a.TryGetSelectedImage().Bytes[0]==9,"floating exact");}});
        Test("Word external picture reports useful error",delegate{var app=Word(7);((FakePicture)app.Selection.InlineShapes[1]).Type=4;using(var a=new WordHostAdapter(app)){Reject(delegate{a.TryGetSelectedImage();});}});
        Test("Word missing exact-range picture uses paragraph ordinal",delegate{
            var app=Word(7);var selected=(FakePicture)app.Selection.InlineShapes[1];
            selected.Range.WordOpenXML=CoreTests.WordWithoutPictures();selected.Range.Start=10;selected.Range.End=11;
            var first=new FakePicture{Type=3,Range=new FakeRange{Start=1,End=2}};
            var paraRange=new FakeRange{WordOpenXML=CoreTests.Word(true).Replace("wp:anchor","wp:inline"),InlineShapes=new FakeCollection(first,selected)};
            selected.Range.Paragraphs=new FakeCollection(new FakeParagraph{Range=paraRange});
            using(var a=new WordHostAdapter(app)){CoreTests.Assert(a.TryGetSelectedImage().Bytes[0]==9,"selected second picture");}
        });
        Test("Word paragraph fallback refuses unknown selected range",delegate{
            var app=Word(7);var selected=(FakePicture)app.Selection.InlineShapes[1];
            selected.Range.WordOpenXML=CoreTests.WordWithoutPictures();selected.Range.Start=10;selected.Range.End=11;
            selected.Range.Paragraphs=new FakeCollection(new FakeParagraph{Range=new FakeRange{WordOpenXML=CoreTests.Word(false).Replace("wp:anchor","wp:inline"),InlineShapes=new FakeCollection(new FakePicture{Type=3,Range=new FakeRange{Start=1,End=2}})}});
            using(var a=new WordHostAdapter(app)){Reject(delegate{a.TryGetSelectedImage();});}
        });
        Test("PPT normal selection saves copy and cleans up",delegate{WithPpt(delegate(FakePpt p){using(var a=new PowerPointHostAdapter(p)){CoreTests.Assert(a.TryGetSelectedImage().Bytes[0]==9,"selected pixels");}CoreTests.Assert(p.ActiveWindow.Presentation.Format==24,"pptx");Clean(p);});});
        Test("PPT macro deck keeps macro-enabled copy",delegate{WithPpt(delegate(FakePpt p){p.ActiveWindow.Presentation.HasVBProject=-1;using(var a=new PowerPointHostAdapter(p)){a.TryGetSelectedImage();}CoreTests.Assert(p.ActiveWindow.Presentation.Format==25,"pptm");Clean(p);});});
        Test("PPT event selection overrides stale selection",delegate{WithPpt(delegate(FakePpt p){var eventSelection=p.ActiveWindow.Selection;p.ActiveWindow.Selection=new FakePptSelection{Type=0};using(var a=new PowerPointHostAdapter(p)){CoreTests.Assert(a.TryGetSelectedImage(eventSelection).Bytes[0]==9,"event image");}Clean(p);});});
        Test("PPT grouped child passes through",delegate{WithPpt(delegate(FakePpt p){p.ActiveWindow.Selection.HasChildShapeRange=-1;using(var a=new PowerPointHostAdapter(p)){CoreTests.Assert(a.TryGetSelectedImage()==null,"group");}CoreTests.Assert(p.ActiveWindow.Presentation.LastPath==null,"no export");});});
        Test("PPT unsupported view passes through",delegate{WithPpt(delegate(FakePpt p){p.ActiveWindow.ViewType=7;using(var a=new PowerPointHostAdapter(p)){CoreTests.Assert(a.TryGetSelectedImage()==null,"view");}});});
        Test("PPT missing image cleans temporary copy",delegate{WithPpt(delegate(FakePpt p){((FakePicture)p.ActiveWindow.Selection.ShapeRange[1]).Id=999;using(var a=new PowerPointHostAdapter(p)){Reject(delegate{a.TryGetSelectedImage();});}Clean(p);});});
        Console.WriteLine("Host adapter tests: "+passed+" passed, "+failed+" failed (simulated Office objects)");
        return failed==0?0:1;
    }
    static void Test(string n,Action f){try{f();passed++;Console.WriteLine("PASS "+n);}catch(Exception e){failed++;Console.WriteLine("FAIL "+n+": "+e);}}
    static void Reject(Action f){try{f();}catch(ImageReadException){return;}throw new Exception("expected image error");}
    static FakeWord Word(int type){
        return new FakeWord{Selection=new FakeWordSelection{
            Type=type, Document=new FakeDocument(),
            InlineShapes=new FakeCollection(new FakePicture{Type=3,Range=new FakeRange{WordOpenXML=CoreTests.Word(false)}}),
            ShapeRange=new FakeCollection(new FakePicture{Type=13,ID=9,Name="Picture 9",Anchor=new FakeAnchor{Paragraphs=new FakeCollection(new FakeParagraph{Range=new FakeRange{WordOpenXML=CoreTests.Word(true)}})}})
        }};
    }
    static void WithPpt(Action<FakePpt> f){CoreTests.WithDeck("normal",delegate(string path){
        f(new FakePpt{ActiveWindow=new FakePptWindow{Presentation=new FakePresentation{Source=path},Selection=new FakePptSelection{
            Type=2,ShapeRange=new FakeCollection(new FakePicture{Type=13,Id=9}),SlideRange=new FakeCollection(new FakeSlide{SlideID=301})
        }}});
    });}
    static void Clean(FakePpt p){string path=p.ActiveWindow.Presentation.LastPath;CoreTests.Assert(path!=null&&!File.Exists(path)&&!Directory.Exists(Path.GetDirectoryName(path)),"temporary folder cleaned");}
}
public class FakeCollection {readonly object[] values;public FakeCollection(params object[] v){values=v;}public int Count{get{return values.Length;}}public object this[int n]{get{return values[n-1];}}}
public class FakeWord {public FakeCollection Documents=new FakeCollection(new object());public FakeWordSelection Selection;public FakeWordWindow ActiveWindow=new FakeWordWindow();}
public class FakeWordWindow{public int Hwnd=0;}
public class FakeWordSelection{public int Type;public FakeDocument Document;public FakeCollection InlineShapes,ShapeRange;}
public class FakeDocument{public string Name="Test.docx";}
public class FakeRange{public string WordOpenXML;public int Start,End;public FakeCollection Paragraphs,InlineShapes;}
public class FakeAnchor{public FakeCollection Paragraphs;}
public class FakeParagraph{public FakeRange Range;}
public class FakePicture{public int Type,ID,Id;public string Name;public FakeRange Range;public FakeAnchor Anchor;}
public class FakePpt{public int HWND=0;public FakeCollection Presentations=new FakeCollection(new object());public FakePptWindow ActiveWindow;}
public class FakePptWindow{public int ViewType=9;public FakePptSelection Selection;public FakePresentation Presentation;}
public class FakePptSelection{public int Type,HasChildShapeRange;public FakeCollection ShapeRange,SlideRange;}
public class FakeSlide{public int SlideID;}
public class FakePresentation{
    public string Name="Test.pptx",Source,LastPath;public int HasVBProject,Format;
    public void SaveCopyAs(string path,int format,int fonts){LastPath=path;Format=format;File.Copy(Source,path);}
}


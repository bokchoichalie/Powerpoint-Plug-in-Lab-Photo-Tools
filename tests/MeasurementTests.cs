using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Windows.Forms;
using System.Xml;
using LabPhotoTools;

internal static class MeasurementTests
{
    private static int checks;
    private static void Check(bool value,string message){checks++;if(!value)throw new Exception(message);}
    private static void Near(double a,double b,string label){Check(Math.Abs(a-b)<.0001*Math.Max(1,Math.Abs(b)),label+": "+a+" != "+b);}
    private static MeasurementItem Item(string kind,params double[] points)
    {MeasurementItem i=new MeasurementItem {Kind=kind};for(int n=0;n<points.Length;n+=2)i.Points.Add(new MeasurePoint(points[n],points[n+1]));return i;}
    private static void Add(MeasurementDocument d,MeasurementItem i){i.Id=d.NextId++;d.Items.Add(i);}
    private static MeasurementDocument Demo()
    {
        MeasurementDocument d=new MeasurementDocument {Width=1200,Height=800};d.Calibrate(Item("line",850,710,1050,710),100,"µm");
        Add(d,Item("line",120,120,380,160));Add(d,Item("circle",650,230,735,230));Add(d,Item("rect",120,350,370,580));
        Add(d,Item("angle",780,510,890,590,1020,480));Add(d,Item("ellipse",550,470,780,550,620,590));return d;
    }
    private static Bitmap TestImage(int height=800)
    {
        Bitmap image=new Bitmap(1200,height);using(Graphics g=Graphics.FromImage(image))
        {
            g.Clear(Color.FromArgb(54,58,64));using(Brush b=new SolidBrush(Color.FromArgb(172,177,183)))
            {g.FillEllipse(b,60,50,370,205);g.FillEllipse(b,555,135,190,190);g.FillRectangle(b,110,340,270,245);g.FillEllipse(b,530,400,340,240);}
            using(Pen p=new Pen(Color.White,9))g.DrawLine(p,850,710,1050,710);
            using(Font f=new Font("Arial",28,FontStyle.Bold,GraphicsUnit.Pixel))g.DrawString("100 µm",f,Brushes.White,890,733);
        }return image;
    }
    [STAThread] private static int Main(string[] args)
    {
        try
        {
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException);
            Directory.CreateDirectory(args[0]);Geometry();Application.EnableVisualStyles();CanvasAndDialog(args[0]);
            if(args.Contains("--powerpoint"))Integration(args[0]);
            Console.WriteLine("PASS: "+checks+" measurement assertions.");return 0;
        }
        catch(Exception ex){Console.Error.WriteLine(ex);return 1;}
    }
    private static void Geometry()
    {
        MeasurementDocument d=Demo();Near(d.MicronsPerPixel,.5,"line scale");
        Near(MeasurementGeometry.ReferenceLength(Item("gap",0,0,100,0,60,40)),40,"parallel scale");
        Near(MeasurementGeometry.ReferenceLength(Item("circle3",10,0,0,10,-10,0)),20,"diameter scale");
        d.Calibrate(Item("gap",0,0,100,0,60,40),20,"µm");Near(d.MicronsPerPixel,.5,"gap calibrated");
        d.Calibrate(Item("circle3",10,0,0,10,-10,0),.01,"mm");Near(d.MicronsPerPixel,.5,"circle calibrated");
        var r=MeasurementGeometry.Build(Item("line",0,0,3,4),d);Near(r.Length,5,"3-4-5 line");
        Near(MeasurementGeometry.Build(Item("rect",20,30,10,5),d).Area,250,"reverse rectangle");
        Near(MeasurementGeometry.Build(Item("rect3",0,0,3,4,-4,3),d).Area,25,"rotated rectangle");
        Near(MeasurementGeometry.Build(Item("circle3",10,0,0,10,-10,0),d).Area,100*Math.PI,"3-point circle area");
        Near(MeasurementGeometry.Build(Item("ellipse",0,0,20,0,10,5),d).Area,50*Math.PI,"ellipse area");
        Near(MeasurementGeometry.Build(Item("angle",1,0,0,0,0,1),d).Angle,90,"angle");
        MeasurementItem obtuse=Item("hangle",0,0,1,1);obtuse.Obtuse=true;Near(MeasurementGeometry.Build(obtuse,d).Angle,135,"obtuse angle");
        Near(MeasurementGeometry.Build(Item("vangle",0,0,1,1),d).Angle,45,"vertical angle");
        Near(MeasurementGeometry.Build(Item("polygon",0,0,10,0,10,10,0,10),d).Area,100,"polygon area");
        Near(MeasurementGeometry.Build(Item("polyline",0,0,3,4,6,0),d).Length,10,"polyline");
        Check(MeasurementGeometry.Build(Item("curve",0,0,50,100,100,0),d).Area>3000,"sampled quadratic curve area");
        Near(MeasurementGeometry.Build(Item("pointline",0,0,10,0,30,5),d).Length,5,"infinite reference line");
        MeasurementDocument circles=new MeasurementDocument {Width=1000,Height=1000,MicronsPerPixel=1};Add(circles,Item("circle",10,10,15,10));Add(circles,Item("circle",30,10,35,10));
        MeasurementItem distance=new MeasurementItem {Kind="circle_distance",CircleA=1,CircleB=2};Add(circles,distance);
        r=MeasurementGeometry.Build(distance,circles);Near(r.Length,20,"circle centers");Near(r.Width,10,"surface gap");
        circles.Items[1].Points[0].X=40;circles.Items[1].Points[1].X=45;Near(MeasurementGeometry.Build(distance,circles).Length,30,"circle reference updates");
        circles.Delete(new[]{1});Check(circles.Items.Count==1,"delete dependent distance");
        foreach(MeasurementItem invalid in new[]{Item("line",1,1,1,1),Item("circle3",0,0,1,1,2,2),Item("ellipse",0,0,0,0,1,1),Item("rect3",0,0,10,0,5,0)})
        {bool failed=false;try{MeasurementGeometry.Build(invalid,d);}catch(InvalidOperationException){failed=true;}Check(failed,"reject degenerate "+invalid.Kind);}
        MeasurementDocument restored=MeasurementDocument.Deserialize(d.Serialize());Near(restored.MicronsPerPixel,d.MicronsPerPixel,"persistence scale");Check(restored.Items.Count==5,"persistence objects");
        Check(d.Csv().Contains("ID,Tool,Result")&&d.Csv().Contains("mm"),"CSV output");
        var p=PowerPointHost.MeasurementToSlide(new MeasurePoint(0,0),new MeasurementDocument {Width=100,Height=100},new PhotoSnapshot {Left=20,Top=30,Width=100,Height=100,Rotation=90});Near(p.X,120,"rotated point x");Near(p.Y,30,"rotated point y");
        XmlDocument xml=new XmlDocument();xml.LoadXml(new Connect().GetCustomUI("test"));XmlNamespaceManager ns=new XmlNamespaceManager(xml.NameTable);ns.AddNamespace("r",xml.DocumentElement.NamespaceURI);
        Check(xml.SelectSingleNode("//r:tab[@id='labPhotoTab']//r:button[@onAction='OpenMeasurement']",ns)!=null,"measurement inside Lab Photo Tools");
        Check(xml.SelectSingleNode("//r:tab[@id='labMeasurementTab']",ns)==null,"no separate measurement tab");
    }
    private static IEnumerable<Control> Children(Control root){foreach(Control c in root.Controls){yield return c;foreach(Control child in Children(c))yield return child;}}
    private static void Reveal(Control child)
    {for(Control p=child.Parent;p!=null;p=p.Parent){ScrollableControl viewport=p as ScrollableControl;if(viewport!=null&&viewport.AutoScroll)viewport.ScrollControlIntoView(child);}Application.DoEvents();}
    private static void Input(Control c,string method,EventArgs args)
    {c.GetType().GetMethod(method,BindingFlags.Instance|BindingFlags.NonPublic).Invoke(c,new object[]{args});}
    private static void PrecisionKeys()
    {
        using(Form form=new Form {ClientSize=new Size(800,600)})using(Bitmap bitmap=TestImage())
        using(MeasurementCanvas c=new MeasurementCanvas {Source=bitmap,Document=new MeasurementDocument {Width=2400,Height=1600}})
        {
            form.Controls.Add(c);form.Show();c.SetTool("line",true);c.Zoom(4);Application.DoEvents();
            Point start=Point.Round(c.ToScreen(new MeasurePoint(1000,700)));
            Input(c,"OnMouseMove",new MouseEventArgs(MouseButtons.None,0,start.X,start.Y,0));MeasurePoint initial=c.HoverPoint;
            foreach(Keys key in new[]{Keys.D,Keys.S,Keys.A,Keys.W})
            {
                MeasurePoint old=c.HoverPoint;Input(c,"OnKeyDown",new KeyEventArgs(key));
                Near(c.HoverPoint.X-old.X,key==Keys.D?2:key==Keys.A?-2:0,"WASD loaded pixel X "+key);
                Near(c.HoverPoint.Y-old.Y,key==Keys.S?2:key==Keys.W?-2:0,"WASD loaded pixel Y "+key);
                Point cursor=c.PointToClient(Cursor.Position);Check(cursor==Point.Round(c.ToScreen(c.HoverPoint)),"native cursor follows "+key);
                MeasurePoint exact=c.HoverPoint;Input(c,"OnMouseMove",new MouseEventArgs(MouseButtons.None,0,cursor.X,cursor.Y,0));Near(c.HoverPoint.X,exact.X,"synthetic mouse preserves subpixel");
            }
            Near(c.HoverPoint.X,initial.X,"WASD round trip");
            Input(c,"OnKeyDown",new KeyEventArgs(Keys.D));MeasurePoint first=c.HoverPoint;
            Point click=c.PointToClient(Cursor.Position);Input(c,"OnMouseDown",new MouseEventArgs(MouseButtons.Left,1,click.X,click.Y,0));
            MeasurementItem reference=null;c.CalibrationReady+=delegate(MeasurementItem i){reference=i;};
            Input(c,"OnKeyDown",new KeyEventArgs(Keys.D|Keys.Shift));Input(c,"OnKeyDown",new KeyEventArgs(Keys.Enter));
            Check(reference!=null,"keyboard and mouse commit calibration");Near(reference.Points[0].X,first.X,"click keeps precision coordinate");Near(MeasurementGeometry.ReferenceLength(reference),20,"Shift ten loaded pixels");
            c.Rotation=90;c.SetTool("line",true);MeasurePoint before=c.HoverPoint;Input(c,"OnKeyDown",new KeyEventArgs(Keys.D));Near(c.HoverPoint.X,before.X,"rotated nudge X");Near(c.HoverPoint.Y,before.Y-2,"rotated nudge follows visible right");
            form.Close();
        }
    }
    private static void CanvasAndDialog(string output)
    {
        Console.WriteLine("Precision input");
        PrecisionKeys();
        using(MeasurementCanvas canvas=new MeasurementCanvas())
        {
            canvas.Document=new MeasurementDocument {Width=1000,Height=600};canvas.Source=new Bitmap(1000,600);canvas.Size=new Size(700,500);
            canvas.SetTool("line",false);canvas.ClickImage(new MeasurePoint(100,100));canvas.ClickImage(new MeasurePoint(200,100));Check(canvas.Document.Items.Count==0,"uncalibrated measurement blocked");
            canvas.Document.Calibrate(Item("line",0,0,100,0),10,"µm");canvas.ClickImage(new MeasurePoint(100,100));canvas.ClickImage(new MeasurePoint(200,100));Check(canvas.Document.Items.Count==1,"two-click line created");
            canvas.Undo();Check(canvas.Document.Items.Count==0,"undo add");canvas.Redo();Check(canvas.Document.Items.Count==1,"redo add");
            canvas.SetTool("polygon",false);canvas.ClickImage(new MeasurePoint(10,10));canvas.ClickImage(new MeasurePoint(100,10));canvas.ClickImage(new MeasurePoint(100,100));canvas.Finish();Check(canvas.Document.Items.Count==2,"finish polygon");
            MeasurePoint point=new MeasurePoint(140,70);var a=canvas.ToImage(canvas.ToScreen(point));Near(a.X,140,"canvas coordinates");canvas.Zoom(3);a=canvas.ToImage(canvas.ToScreen(point));Near(a.Y,70,"zoom coordinates");
            canvas.Source.Dispose();
        }
        using(MeasurementForm form=new MeasurementForm(new PowerPointHost(new object()),new SelectionSnapshot(),new MeasurementSession {Image=TestImage(),Document=Demo()}))
        {
            Console.WriteLine("Base layout");
            form.Show();Application.DoEvents();Rectangle work=Screen.FromHandle(form.Handle).WorkingArea;Check(work.Contains(form.Bounds)&&form.Width>=work.Width*.94&&form.Height>=work.Height*.94,"dialog fills monitor work area");
            Check(form.Canvas.Focused,"measurement keys ready when dialog opens");form.ClientSize=new Size(1280,830);Application.DoEvents();
            Check(form.FormBorderStyle==FormBorderStyle.FixedDialog&&!form.MaximizeBox,"fixed dialog border");
            using(Bitmap shot=new Bitmap(form.Width,form.Height)){form.DrawToBitmap(shot,new Rectangle(Point.Empty,shot.Size));shot.Save(Path.Combine(output,"measurement-dialog-wide.png"));}
            foreach(Size size in new[]{new Size(780,520),new Size(1024,680),new Size(1280,830),new Size(360,340)})
            {
                form.MinimumSize=Size.Empty;form.ClientSize=size;Application.DoEvents();
                foreach(TableLayoutPanel panel in Children(form).OfType<TableLayoutPanel>())
                {var children=panel.Controls.Cast<Control>().Where(c=>c.Visible).ToList();for(int i=0;i<children.Count;i++)for(int j=i+1;j<children.Count;j++)Check(!children[i].Bounds.IntersectsWith(children[j].Bounds),"dialog overlaps at "+size+" "+children[i].Text+" / "+children[j].Text);}
                Check(form.Canvas.Width>120&&form.Canvas.Height>80,"canvas visible at "+size+": "+form.Canvas.Size+" body="+form.Canvas.Parent.Size);
                Check(form.Canvas.Top<=16&&form.Canvas.Height>=form.ClientSize.Height-40,"photo uses full window height "+size);
                Rectangle imageBounds=form.Canvas.RectangleToScreen(form.Canvas.ClientRectangle);
                foreach(Control button in Children(form).OfType<Button>())
                {Rectangle bounds=button.RectangleToScreen(button.ClientRectangle);for(Control p=button.Parent;p!=form;p=p.Parent)bounds=Rectangle.Intersect(bounds,p.RectangleToScreen(p.ClientRectangle));if(!bounds.IsEmpty)Check(bounds.Left>=imageBounds.Right,"all buttons to the right "+button.Text+" at "+size);}
                Button apply=Children(form).OfType<Button>().Single(b=>b.Text=="측정 사진 복사");Reveal(apply);Rectangle bnd=form.RectangleToClient(apply.RectangleToScreen(apply.ClientRectangle));Check(form.ClientRectangle.Contains(bnd),"apply accessible "+size+" bounds="+bnd+" root="+form.Controls[0].Controls[0].Bounds);
                Control loupe=Children(form).Single(c=>c.Name=="MeasurementMagnifier");Near((double)loupe.Width/loupe.Height,4.0/3,"4:3 magnifier");
                Control unit=Children(form).Single(c=>c.AccessibleName=="스케일 단위");Check(unit.Right<=unit.Parent.ClientSize.Width&&unit.Bottom<=unit.Parent.ClientSize.Height,"scale unit not clipped");Check(loupe.PointToScreen(Point.Empty).X>unit.PointToScreen(Point.Empty).X+unit.Width,"magnifier right of scale values");
            }
            form.ClientSize=new Size(780,520);Application.DoEvents();using(Bitmap shot=new Bitmap(form.Width,form.Height)){form.DrawToBitmap(shot,new Rectangle(Point.Empty,shot.Size));shot.Save(Path.Combine(output,"measurement-dialog-small.png"));}form.Close();
        }
        foreach(float scale in new[]{1.25f,1.5f,2f,3f})
        using(MeasurementForm form=new MeasurementForm(new PowerPointHost(new object()),new SelectionSnapshot(),new MeasurementSession {Image=TestImage(),Document=Demo()}))
        {
            Console.WriteLine("DPI layout "+scale);
            form.Show();Application.DoEvents();form.AutoScaleMode=AutoScaleMode.None;form.Scale(new SizeF(scale,scale));form.Font=new Font("맑은 고딕",9.5f*scale);
            foreach(Control child in Children(form))child.Font=new Font(child.Font.FontFamily,9.5f*scale,child.Font.Style);
            foreach(Size size in new[]{new Size(780,520),new Size(1340,700),new Size(1900,1000)})
            {
                form.MinimumSize=Size.Empty;form.ClientSize=size;Application.DoEvents();
                foreach(TableLayoutPanel panel in Children(form).OfType<TableLayoutPanel>())
                {var controls=panel.Controls.Cast<Control>().Where(c=>c.Visible).ToList();for(int i=0;i<controls.Count;i++)for(int j=i+1;j<controls.Count;j++)Check(!controls[i].Bounds.IntersectsWith(controls[j].Bounds),"DPI overlap "+scale+" at "+size+" "+controls[i].GetType().Name+" "+controls[i].Text+" "+controls[i].Bounds+" / "+controls[j].GetType().Name+" "+controls[j].Text+" "+controls[j].Bounds);}
                Button apply=Children(form).OfType<Button>().Single(b=>b.Text=="측정 사진 복사");
                Reveal(apply);
                Check(form.ClientRectangle.Contains(form.RectangleToClient(apply.RectangleToScreen(apply.ClientRectangle))),"DPI apply reachable by scroll "+scale+" at "+size);
                Check(form.Canvas.Height>=45,"DPI canvas remains visible "+scale+" at "+size);
                Check(form.Canvas.Height>=form.ClientSize.Height-64*scale,"DPI image keeps full height "+scale+" at "+size+" canvas="+form.Canvas.Bounds+" parent="+form.Canvas.Parent.Bounds+" viewport="+form.Controls[0].Bounds);
                Check(!((Panel)form.Controls[0]).AutoScroll,"only sidebar scrolls "+scale);
                Control unit=Children(form).Single(c=>c.AccessibleName=="스케일 단위");Check(unit.Right<=unit.Parent.ClientSize.Width&&unit.Bottom<=unit.Parent.ClientSize.Height,"DPI scale unit not clipped "+scale);
                Control loupe=Children(form).Single(c=>c.Name=="MeasurementMagnifier");Near((double)loupe.Width/loupe.Height,4.0/3,"DPI 4:3 magnifier");
            }
            form.Close();
        }
        using(MeasurementForm form=new MeasurementForm(new PowerPointHost(new object()),new SelectionSnapshot(),new MeasurementSession {Image=TestImage(900),Document=new MeasurementDocument {Width=1200,Height=900}}))
        {
            form.Show();Application.DoEvents();
            form.ClientSize=new Size(3650,1860);Application.DoEvents();
            double width=1200*form.Canvas.ViewScale,height=900*form.Canvas.ViewScale;
            File.WriteAllText(Path.Combine(output,"measurement-photo-size.txt"),"Client="+form.ClientSize+" Canvas="+form.Canvas.Size+" Photo="+width.ToString("0")+"x"+height.ToString("0")+" ComparedToScreenshot="+(width/1407).ToString("0.000")+"x width / "+(width*height/(1407*1055.0)).ToString("0.000")+"x area");
            Check(width>1407*1.65&&height>1055*1.65,"4:3 photo at least 65% larger than supplied screenshot");
            using(Bitmap shot=new Bitmap(form.Width,form.Height)){form.DrawToBitmap(shot,new Rectangle(Point.Empty,shot.Size));shot.Save(Path.Combine(output,"measurement-dialog-large-dpi.png"));}
            form.Close();
        }
    }
    private static PhotoSnapshot Snapshot(dynamic s)
    {return (PhotoSnapshot)typeof(PowerPointHost).GetMethod("SnapshotPhoto",BindingFlags.Static|BindingFlags.NonPublic).Invoke(null,new object[]{(object)s});}
    private static void Similar(Color a,Color b,string name)
    {Check(Math.Abs(a.R-b.R)+Math.Abs(a.G-b.G)+Math.Abs(a.B-b.B)<35,name+": "+a+" / "+b);}
    private static void Orientation(PowerPointHost host,dynamic app,dynamic presentation,string output)
    {
        string path=Path.Combine(output,"measurement-orientation-source.png");
        using(Bitmap image=new Bitmap(600,400))using(Graphics g=Graphics.FromImage(image))
        {g.FillRectangle(Brushes.Red,0,0,300,200);g.FillRectangle(Brushes.Lime,300,0,300,200);g.FillRectangle(Brushes.Blue,0,200,300,200);g.FillRectangle(Brushes.Yellow,300,200,300,200);image.Save(path);}
        foreach(string mode in new[]{"normal","h","v","hv","r90","r180","h-r27","nested"})
        {
            dynamic slide=presentation.Slides.Add((int)presentation.Slides.Count+1,12);app.ActiveWindow.View.GotoSlide((int)slide.SlideIndex);
            dynamic photo=slide.Shapes.AddPicture(path,0,-1,280f,220f,400f,266.6667f);photo.Name="Orientation_"+mode;
            if(mode=="h"||mode=="hv"||mode=="h-r27")photo.Flip(0);if(mode=="v"||mode=="hv")photo.Flip(1);
            if(mode=="r90")photo.Rotation=90f;if(mode=="r180")photo.Rotation=180f;if(mode=="h-r27")photo.Rotation=27f;
            dynamic originalGroup=null;
            if(mode=="nested")
            {
                dynamic text=slide.Shapes.AddTextbox(1,220f,150f,150f,30f);text.TextFrame.TextRange.Text="Do not duplicate me";
                dynamic other=slide.Shapes.AddPicture(path,0,-1,40f,50f,60f,40f);
                originalGroup=slide.Shapes.Range(new object[]{(string)photo.Name,(string)text.Name,(string)other.Name}).Group();originalGroup.Rotation=27f;originalGroup.Flip(1);
            }
            SelectionSnapshot selected=new SelectionSnapshot {Presentation=presentation,Slide=slide,SlideWidth=960,SlideHeight=720};selected.Photos.Add(Snapshot(photo));
            using(MeasurementSession session=host.CreateMeasurementSession(selected))
            {
                var d=session.Document;d.Calibrate(Item("line",d.Width*.2,d.Height*.3,d.Width*.4,d.Height*.3),100,"µm");Add(d,Item("line",d.Width*.2,d.Height*.3,d.Width*.4,d.Height*.3));
                int originalCount=(int)slide.Shapes.Count;dynamic saved=host.ApplyMeasurements(selected,d);Check((int)slide.Shapes.Count==originalCount+1,"one copied result "+mode);
                Check((int)saved.GroupItems.Count==3,"only photo, measurement line and result text copied "+mode);
                saved.Select(-1);var copied=host.ReadMeasurementSelection();
                using(MeasurementSession reopened=host.CreateMeasurementSession(copied))
                {Near(reopened.Rotation,session.Rotation,"rotation retained on save "+mode);Similar(reopened.Image.GetPixel(reopened.Image.Width/4,reopened.Image.Height/4),session.Image.GetPixel(session.Image.Width/4,session.Image.Height/4),"no repeated flip "+mode);}
                PhotoSnapshot copiedPhoto=copied.Photos[0];MeasurePoint expected=PowerPointHost.MeasurementToSlide(d.Items[0].Points[0],d,copiedPhoto);
                dynamic line=null;for(int n=1;n<=(int)saved.GroupItems.Count;n++){dynamic child=saved.GroupItems.Item(n);if((int)child.Type==5)line=child;}
                Check(line!=null,"editable freeform exists "+mode);Array node=(Array)line.Nodes.Item(1).Points;
                Near(Convert.ToDouble(node.GetValue(node.GetLowerBound(0),node.GetLowerBound(1))),expected.X,"saved line X "+mode);Near(Convert.ToDouble(node.GetValue(node.GetLowerBound(0),node.GetLowerBound(1)+1)),expected.Y,"saved line Y "+mode);
                saved.Delete();
                string rendered=Path.Combine(output,"measurement-orientation-"+mode+".png");slide.Export(rendered,"PNG",960,720);
                if(originalGroup!=null)originalGroup.Ungroup();PhotoSnapshot flattened=Snapshot(photo);
                using(Bitmap slideImage=new Bitmap(rendered))using(MeasurementCanvas c=new MeasurementCanvas {Document=d,Source=session.Image,Rotation=session.Rotation,Size=new Size(840,640),Tool="select"})
                using(Bitmap preview=new Bitmap(840,640))
                {
                    d.Items.Clear();c.DrawToBitmap(preview,new Rectangle(Point.Empty,preview.Size));
                    foreach(double x in new[]{.25,.75})foreach(double y in new[]{.25,.75})
                    {
                        MeasurePoint point=new MeasurePoint(d.Width*x,d.Height*y);PointF view=c.ToScreen(point);MeasurePoint actual=PowerPointHost.MeasurementToSlide(point,d,flattened);
                        Similar(preview.GetPixel((int)view.X,(int)view.Y),slideImage.GetPixel((int)actual.X,(int)actual.Y),"slide and dialog orientation "+mode+" "+x+","+y);
                    }
                    if(mode=="r180")preview.Save(Path.Combine(output,"measurement-orientation-preview.png"));
                }
            }
        }
    }
    private static void Integration(string output)
    {
        dynamic app=null,presentation=null;
        try
        {
            string path=Path.Combine(output,"measurement-test-image.png");using(Bitmap image=TestImage())image.Save(path);
            app=Activator.CreateInstance(Type.GetTypeFromProgID("PowerPoint.Application"));presentation=app.Presentations.Add(-1);presentation.PageSetup.SlideWidth=960f;presentation.PageSetup.SlideHeight=720f;
            dynamic slide=presentation.Slides.Add(1,12);app.ActiveWindow.View.GotoSlide(1);
            dynamic photo=slide.Shapes.AddPicture(path,0,-1,50f,70f,800f,533.3333f);photo.Name="Microscope";
            PowerPointHost host=new PowerPointHost(app);photo.Select(-1);SelectionSnapshot selected=host.ReadMeasurementSelection();
            using(MeasurementSession session=host.CreateMeasurementSession(selected))
            {
                Check(session.Image.Width>=1200,"high resolution preview");Near(session.Document.Width/session.Document.Height,1.5,"preview aspect");
                double s=session.Document.Width/1200;MeasurementDocument doc=Demo();doc.Width=session.Document.Width;doc.Height=session.Document.Height;
                foreach(MeasurementItem i in doc.Items)foreach(MeasurePoint point in i.Points){point.X*=s;point.Y*=s;}
                doc.Calibrate(Item("line",850*s,710*s,1050*s,710*s),100,"µm");doc.PhotoSignature=session.Document.PhotoSignature;
                dynamic unrelated=slide.Shapes.AddTextbox(1,10f,10f,200f,24f);unrelated.TextFrame.TextRange.Text="Unrelated text remains once";
                dynamic copy=host.ApplyMeasurements(selected,doc);Check((int)presentation.Slides.Count==1,"copy photo without adding a slide");Check((int)slide.Shapes.Count==3,"only one photo group added");
                Check((int)copy.Type==6,"native annotations grouped with photo");Near((double)photo.Left,50,"original photo position");Check(PowerPointHost.ReadMeasurementData(photo)==null,"original photo tags unchanged");Check((string)unrelated.TextFrame.TextRange.Text=="Unrelated text remains once","other text unchanged");
                slide.Export(Path.Combine(output,"measurement-slide.png"),"PNG",1440,1080);
                dynamic group=copy;group.Select(-1);selected=host.ReadMeasurementSelection();
                MeasurementDocument read=PowerPointHost.ReadMeasurementData(selected.Photos[0].Shape);Near(read.MicronsPerPixel,doc.MicronsPerPixel,"tag calibration");Check(read.Items.Count==5,"tag objects");
                using(MeasurementSession again=host.CreateMeasurementSession(selected))Check(again.Document.Items.Count==5,"reopen labelled group");
                photo.Delete();dynamic number=host.AddNumberLabel("square",2);Check((string)number.TextFrame2.TextRange.Text=="2","number measured photo");
                dynamic parent=PowerPointHost.TopMeasurementParent(selected.Photos[0].Shape);parent.Select(-1);
                SelectionSnapshot before=host.ReadMeasurementSelection();double old=before.Photos[0].Width;host.ArrangeAllPhotos();
                parent.Select(-1);selected=host.ReadMeasurementSelection();Check(selected.Photos[0].Width<old,"smart arrange scaled measured group");
                using(MeasurementSession resized=host.CreateMeasurementSession(selected)){Near(resized.Document.MicronsPerPixel,doc.MicronsPerPixel,"resize preserves scale");Check(resized.Document.Items.Count==5,"resized document");}
                string previousNumberGroup=(string)number.ParentGroup.Name;dynamic repeat=host.ApplyMeasurements(selected,read);Check((int)presentation.Slides.Count==1,"reapply stays on same slide");Check((string)number.ParentGroup.Name==previousNumberGroup,"original number label remains in original group");
                repeat.Select(-1);string repeatName=repeat.Name;SelectionSnapshot latest=host.ReadMeasurementSelection();Check(PowerPointHost.ReadMeasurementData(latest.Photos[0].Shape).Items.Count==5,"reapply no duplicate measurements");
                string saved=Path.Combine(output,"measurement-validation.pptx");presentation.SaveAs(saved,24);presentation.Close();presentation=null;
                presentation=app.Presentations.Open(saved,0,0,-1);app.ActiveWindow.View.GotoSlide(1);presentation.Slides.Item(1).Shapes.Item(repeatName).Select(-1);selected=host.ReadMeasurementSelection();
                Check(PowerPointHost.ReadMeasurementData(selected.Photos[0].Shape).Items.Count==5,"persist after reopen pptx");
                host.ResetMeasurements(selected);Check((int)presentation.Slides.Count==1,"reset copies only photo");
                var resetSelection=host.ReadMeasurementSelection();dynamic resetGroup=PowerPointHost.TopMeasurementParent(resetSelection.Photos[0].Shape);Check(PowerPointHost.ReadMeasurementData(resetSelection.Photos[0].Shape)==null,"reset removes calibration");Check((int)resetGroup.Type==6,"reset keeps number photo grouped");
            }
            dynamic rotated=presentation.Slides.Add(2,12);app.ActiveWindow.View.GotoSlide(2);
            dynamic rp=rotated.Shapes.AddPicture(path,0,-1,220f,180f,420f,280f);rp.Name="RotatedCrop";rp.PictureFormat.CropLeft=25f;rp.Flip(0);rp.Rotation=27f;
            dynamic label=host.AddNumberLabel("circle",4);dynamic ng=label.ParentGroup;ng.Rotation=12f;ng.Select(-1);
            selected=host.ReadMeasurementSelection();
            using(MeasurementSession cropped=host.CreateMeasurementSession(selected))
            {
                var d=cropped.Document;d.Calibrate(Item("line",d.Width*.1,d.Height*.1,d.Width*.3,d.Height*.1),25,"µm");Add(d,Item("line",d.Width*.1,d.Height*.1,d.Width*.3,d.Height*.1));
                dynamic measured=host.ApplyMeasurements(selected,d);measured.Select(-1);selected=host.ReadMeasurementSelection();
                rotated.Export(Path.Combine(output,"measurement-rotated-cropped-slide.png"),"PNG",1440,1080);
                using(MeasurementSession re=host.CreateMeasurementSession(selected))Check(re.Document.Items.Count==1,"cropped flipped rotated labelled group reopens");
                dynamic gp=measured;gp.Width=(float)((double)gp.Width*.8);gp.Height=(float)((double)gp.Height*.8);gp.Select(-1);selected=host.ReadMeasurementSelection();
                using(MeasurementSession re=host.CreateMeasurementSession(selected))Check(re.Document.HasScale,"cropped image resizing preserves calibration");
            }
            dynamic all=presentation.Slides.Add((int)presentation.Slides.Count+1,12);app.ActiveWindow.View.GotoSlide((int)all.SlideIndex);
            dynamic allPhoto=all.Shapes.AddPicture(path,0,-1,30f,50f,800f,533.3333f);allPhoto.Select(-1);selected=host.ReadMeasurementSelection();
            using(MeasurementSession s=host.CreateMeasurementSession(selected))
            {
                var d=s.Document;d.Calibrate(Item("line",0,0,100,0),50,"nm");
                foreach(var i in new[]{Item("line",50,50,150,150),Item("rect",60,60,200,200),Item("rect3",100,20,150,50,150,120),Item("circle",200,200,250,200),Item("circle3",310,300,300,310,290,300),Item("ellipse",300,100,400,100,350,160),Item("angle",400,100,500,200,550,50),Item("hangle",600,100,650,130),Item("vangle",700,100,740,130),Item("polygon",600,300,650,400,500,400),Item("polyline",700,400,720,420,740,400),Item("lasso",600,500,650,550,600,600,550,550),Item("curve",800,500,850,550,900,500),Item("gap",100,500,200,500,150,550,150,575),Item("pointline",400,600,500,600,450,650),Item("point",800,200),Item("text",900,200),Item("draw",900,400,920,410,930,390),Item("arrow",900,500,970,570)})
                {if(i.Kind=="text")i.Note="한글 주석 / µm";Add(d,i);}
                Add(d,new MeasurementItem {Kind="circle_distance",CircleA=4,CircleB=5});
                dynamic result=host.ApplyMeasurements(selected,d);result.Select(-1);var actualSelection=host.ReadMeasurementSelection();
                Check(PowerPointHost.ReadMeasurementData(actualSelection.Photos[0].Shape).Items.Count==20,"all 20 manual tools exported with editable geometry");
            }
            Orientation(host,app,presentation,output);
        }
        finally
        {
            if(presentation!=null){try{presentation.Saved=-1;presentation.Close();}catch{}}
            if(app!=null){try{app.Quit();}catch{}try{Marshal.ReleaseComObject(app);}catch{}}
        }
    }
}

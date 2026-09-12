using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using LabPhotoTools;

internal static class PowerPointIntegration
{
    private sealed class TestHost : PowerPointHost
    {
        internal TestHost(object app) : base(app) { }
        public new object AddNumberLabel(string style, int? number) { return base.AddNumberLabel(style,number,new NumberLabelSettings()); }
    }
    private static int checks;
    private static void Check(bool condition, string message) { checks++; if(!condition) throw new Exception(message); }
    [STAThread]
    private static int Main(string[] args)
    {
        dynamic app=null, presentation=null, originalWindow=null;
        try
        {
            string output=args[0]; Directory.CreateDirectory(output);
            bool skipImageActions=args.Length>1 && args[1]=="--skip-image-actions";
            string imagePath=Path.Combine(output,"sample.png");
            using(Bitmap bitmap=new Bitmap(360,240)) using(Graphics g=Graphics.FromImage(bitmap))
            { g.Clear(Color.White);g.FillEllipse(Brushes.SteelBlue,80,25,185,185);bitmap.Save(imagePath); }
            app=Activator.CreateInstance(Type.GetTypeFromProgID("PowerPoint.Application"));
            try { originalWindow=app.ActiveWindow; } catch { }
            presentation=app.Presentations.Add(-1);
            presentation.PageSetup.SlideWidth=960f; presentation.PageSetup.SlideHeight=540f;
            dynamic slide=presentation.Slides.Add(1,12);
            dynamic title=slide.Shapes.AddTextbox(1,20f,10f,300f,30f); title.TextFrame.TextRange.Text="LabPhotoTools 0.1.11 integration test";
            for(int i=0;i<10;i++)
            {
                dynamic photo=slide.Shapes.AddPicture(imagePath,0,-1,10f+i*20,50f+i*6,100f,70f);
                if(i==3) photo.Rotation=27f;
            }
            app.ActiveWindow.View.GotoSlide(1);
            TestHost host=new TestHost(app);
            SelectionSnapshot before=host.ReadAllPhotos();
            Check(before.Photos.Count==10,"All photos without selection");
            ClearSelection(app);
            host.ArrangeAllPhotos();
            Check((double)title.Left==20 && (double)title.Top==10,"Text should stay put");
            SelectionSnapshot after=host.ReadAllPhotos();
            Check(Math.Abs(after.Photos[3].Rotation-27)<.01,"Rotation preserved");
            ValidatePlan(after);
            slide.Export(Path.Combine(output,"smart-layout.png"),"PNG",1440,810);
            app.CommandBars.ExecuteMso("Undo");
            Check(Math.Abs((double)((dynamic)before.Photos[0].Shape).Left-before.Photos[0].Left)<.05,"Smart layout Undo");
            host.ArrangeAllPhotos();
            Console.WriteLine("Actual PowerPoint: all-photo layout, rotated bounds, text preservation and Undo passed.");

            dynamic numbered=presentation.Slides.Add(2,12); app.ActiveWindow.View.GotoSlide(2);
            // Each label attaches to one photo. The photo grid is deliberately
            // larger than the label so alignment and grouping remain visible.
            for(int i=0;i<52;i++) numbered.Shapes.AddPicture(imagePath,0,-1,20f+(i%13)*70f,45f+(i/13)*105f,54f,76f);
            foreach(string style in new[]{"square","circle","paren","suffix"})
            {
                for(int n=0;n<=10;n++)
                {
                    dynamic shape=host.AddNumberLabel(style,n);
                    Check((string)shape.TextFrame2.TextRange.Text==NumberLabels.Text(style,n),"Number text");
                    if(style=="square" || style=="circle")
                    { Check((int)shape.Fill.ForeColor.RGB==0xFFFFFF,"White fill");Check(Math.Abs((double)shape.Width-(double)shape.Height)<.01,"Square/circle ratio"); }
                    else Check((int)shape.Fill.Visible==0 && (int)shape.Line.Visible==0,"Plain label frame");
                    Check(Math.Abs((double)shape.Width-(double)shape.Height)<.01 && (double)shape.Width>=new NumberLabelSettings().BaseContainerSide-.01,"Every label has a proportional square container");
                    Check((int)shape.TextFrame2.AutoSize==0,"Do not shrink label text or container automatically");
                }
                dynamic next=host.AddNumberLabel(style,null);
                Check((string)next.TextFrame2.TextRange.Text==NumberLabels.Text(style,11),"Next 11");
                next=host.AddNumberLabel(style,null);
                Check((string)next.TextFrame2.TextRange.Text==NumberLabels.Text(style,12),"Next 12");
            }
            Check(CountNumberLabels(numbered.Shapes)==52,"All number labels");
            Check(CountAttachedPhotos(numbered.Shapes)==52,"Every label attaches to one photo");
            numbered.Export(Path.Combine(output,"number-labels.png"),"PNG",1440,810);
            app.CommandBars.ExecuteMso("Undo");
            Check(CountNumberLabels(numbered.Shapes)==51 && CountAttachedPhotos(numbered.Shapes)==51,"Number Undo");
            dynamic restored=host.AddNumberLabel("suffix",null);
            Check((string)restored.TextFrame2.TextRange.Text=="12)","Undo-aware sequence");
            numbered.Shapes.AddPicture(imagePath,0,-1,900f,450f,40f,55f);
            Console.WriteLine("Actual PowerPoint: 44 presets, next values, labels attached to photographs and Undo passed.");

            dynamic groups=presentation.Slides.Add(3,12); app.ActiveWindow.View.GotoSlide(3);
            dynamic ga=groups.Shapes.AddPicture(imagePath,0,-1,20f,20f,100f,80f);
            dynamic gb=groups.Shapes.AddPicture(imagePath,0,-1,140f,20f,90f,80f);
            dynamic group=groups.Shapes.Range(new object[]{(string)ga.Name,(string)gb.Name}).Group();
            groups.Shapes.AddPicture(imagePath,0,-1,300f,70f,80f,140f);
            host.ArrangeAllPhotos();
            SelectionSnapshot grouped=host.ReadAllPhotos();
            Check(grouped.Photos.Count==3,"Group photos included"); ValidatePlan(grouped);
            Check((int)group.Type==6,"Group preserved");
            Console.WriteLine("Actual PowerPoint: grouped photographs passed.");

            TestMixedSelection(app,presentation,host,imagePath);

            if(!skipImageActions)
            {
                int slideCountBeforeImages=(int)presentation.Slides.Count;
                dynamic imageSlide=presentation.Slides.Add(slideCountBeforeImages+1,12); app.ActiveWindow.View.GotoSlide((int)imageSlide.SlideIndex);
                dynamic imageShape=imageSlide.Shapes.AddPicture(imagePath,0,-1,100f,100f,180f,120f); imageShape.Select(-1);
                SelectionSnapshot selected=host.ReadSelection();
                host.ApplyImagesAsync(selected,false,true,2,CancellationToken.None).GetAwaiter().GetResult();
                Check((int)presentation.Slides.Count==slideCountBeforeImages+2,"Rotation duplicates slide");
                Check(Math.Abs((double)imageShape.Rotation)<.001,"Original photo unchanged");
                app.ActiveWindow.View.GotoSlide((int)imageSlide.SlideIndex); imageShape.Select(-1);
                host.ApplyImagesAsync(host.ReadSelection(),true,false,0,CancellationToken.None).GetAwaiter().GetResult();
                Check((int)presentation.Slides.Count==slideCountBeforeImages+3,"Background duplicates slide");
                Console.WriteLine("Actual PowerPoint: rotation and offline background removal passed.");
            }
            TestPhotoLabelAttachment(app,presentation,host,imagePath);
            TestNumberFormatting(app,presentation,host,output,imagePath);
            presentation.SaveAs(Path.Combine(output,"LabPhotoTools-0.1.11-validation.pptx"),24);
            // Re-open verifies that the sequence survives saving and closing.
            presentation.Close(); presentation=null;
            presentation=app.Presentations.Open(Path.Combine(output,"LabPhotoTools-0.1.11-validation.pptx"),0,0,-1);
            app.ActiveWindow.View.GotoSlide(2);
            dynamic persisted=host.AddNumberLabel("square",null);
            Check((string)persisted.TextFrame2.TextRange.Text=="13","Sequence persists after reopening");
            presentation.Saved=-1;
            Console.WriteLine("PASS: "+checks+" PowerPoint integration assertions.");
            return 0;
        }
        catch(Exception ex) { Console.Error.WriteLine(ex); return 1; }
        finally
        {
            if(presentation!=null) { try { presentation.Saved=-1;presentation.Close(); } catch { } }
            if(originalWindow!=null) { try { originalWindow.Activate(); } catch { } }
            // Never close an existing PowerPoint instance or a user document.
        }
    }
    private static void TestMixedSelection(dynamic app, dynamic presentation, PowerPointHost host, string imagePath)
    {
        dynamic manual=presentation.Slides.Add((int)presentation.Slides.Count+1,12); app.ActiveWindow.View.GotoSlide((int)manual.SlideIndex);
        dynamic first=manual.Shapes.AddPicture(imagePath,0,-1,50f,100f,110f,70f);
        dynamic text=manual.Shapes.AddTextbox(1,25f,18f,260f,30f); text.TextFrame.TextRange.Text="Do not move this text";
        dynamic second=manual.Shapes.AddPicture(imagePath,0,-1,220f,115f,95f,120f);
        dynamic box=manual.Shapes.AddShape(1,410f,90f,80f,55f);
        dynamic untouched=manual.Shapes.AddPicture(imagePath,0,-1,650f,320f,80f,60f);
        RectangleF textBefore=Bounds(text), boxBefore=Bounds(box), untouchedBefore=Bounds(untouched);
        SelectShapes(manual,first,text,second,box);
        SelectionSnapshot selected=host.ReadSelection();
        Check(selected.Photos.Count==2,"Mixed selection must retain only photos");
        Check(selected.Photos.All(p=>(int)p.Id==(int)first.Id || (int)p.Id==(int)second.Id),"Mixed selection included a non-photo");
        host.ApplyLayout(selected,Layout.Arrange(selected.Boxes(),2,12,8,false,0,selected.SlideWidth,selected.SlideHeight));
        CheckUnchanged(text,textBefore,"Manual layout moved text"); CheckUnchanged(box,boxBefore,"Manual layout moved shape"); CheckUnchanged(untouched,untouchedBefore,"Manual layout moved unselected photo");
        SelectShapes(manual,first,text,second,box);
        selected=host.ReadSelection();
        host.ApplyLayout(selected,Layout.Space(selected.Boxes(),18,14,true,true,selected.SlideWidth,selected.SlideHeight));
        CheckUnchanged(text,textBefore,"Spacing moved text"); CheckUnchanged(box,boxBefore,"Spacing moved shape"); CheckUnchanged(untouched,untouchedBefore,"Spacing moved unselected photo");

        dynamic smart=presentation.Slides.Add((int)presentation.Slides.Count+1,12); app.ActiveWindow.View.GotoSlide((int)smart.SlideIndex);
        dynamic smartFirst=smart.Shapes.AddPicture(imagePath,0,-1,15f,75f,100f,70f);
        dynamic smartText=smart.Shapes.AddTextbox(1,35f,15f,280f,30f); smartText.TextFrame.TextRange.Text="Keep smart-layout text";
        dynamic smartSecond=smart.Shapes.AddPicture(imagePath,0,-1,160f,80f,90f,120f);
        dynamic smartBox=smart.Shapes.AddShape(9,400f,35f,60f,60f);
        dynamic smartUntouched=smart.Shapes.AddPicture(imagePath,0,-1,700f,300f,120f,80f);
        RectangleF smartFirstBefore=Bounds(smartFirst), smartSecondBefore=Bounds(smartSecond), smartTextBefore=Bounds(smartText), smartBoxBefore=Bounds(smartBox), smartUntouchedBefore=Bounds(smartUntouched);
        SelectShapes(smart,smartFirst,smartText,smartSecond,smartBox);
        host.ArrangeAllPhotos();
        Check(!SameBounds(smartFirst,smartFirstBefore) || !SameBounds(smartSecond,smartSecondBefore),"Smart layout did not arrange selected photos");
        CheckUnchanged(smartText,smartTextBefore,"Smart layout moved text"); CheckUnchanged(smartBox,smartBoxBefore,"Smart layout moved shape"); CheckUnchanged(smartUntouched,smartUntouchedBefore,"Smart layout moved unselected photo");
        Console.WriteLine("Actual PowerPoint: mixed selections arrange only selected photographs passed.");
    }
    private static void TestProportionalNumberContainers(dynamic app, dynamic presentation, PowerPointHost host, string imagePath)
    {
        dynamic slide=presentation.Slides.Add((int)presentation.Slides.Count+1,12); app.ActiveWindow.View.GotoSlide((int)slide.SlideIndex);
        for(int i=0;i<8;i++) slide.Shapes.AddPicture(imagePath,0,-1,30f+i*110f,60f,90f,70f);
        Dictionary<string,double> seventeenPointSides=new Dictionary<string,double>();
        foreach(float size in new[]{17f,34f})
        {
            NumberLabelSettings settings=new NumberLabelSettings {FontSize=size};
            double expected=settings.BaseContainerSide;
            foreach(string style in new[]{"square","circle","paren","suffix"})
            {
                dynamic shape=host.AddNumberLabel(style,1,settings);
                Check(Math.Abs((double)shape.Width-(double)shape.Height)<.01,"Proportional label must stay square");
                Check((double)shape.Width>=expected-.01,"Every label starts from the 17 pt / 0.85 cm base ratio");
                if(size==17) seventeenPointSides.Add(style,(double)shape.Width);
                else Check(Math.Abs((double)shape.Width-seventeenPointSides[style]*2)<.5,"Container side must double when point size doubles");
                if(style=="square" || style=="circle") Check(Math.Abs((double)shape.Width-expected)<.5,"17 pt square and circle use an 0.85 cm side/diameter");
            }
        }
        Console.WriteLine("Actual PowerPoint: 17 pt to 0.85 cm proportional label containers passed.");
    }
    private static void ClearSelection(dynamic app)
    {
        try { app.ActiveWindow.Selection.Unselect(); } catch { }
    }
    private static void SelectShapes(dynamic slide, params dynamic[] shapes)
    {
        object[] names=new object[shapes.Length];
        for(int i=0;i<shapes.Length;i++) names[i]=(string)shapes[i].Name;
        slide.Shapes.Range(names).Select();
    }
    private static RectangleF Bounds(dynamic shape) { return new RectangleF((float)shape.Left,(float)shape.Top,(float)shape.Width,(float)shape.Height); }
    private static string TagValue(dynamic shape, string tag)
    {
        try { return Convert.ToString(shape.Tags.Item(tag)) ?? ""; } catch { return ""; }
    }
    private static int CountTagged(dynamic shapes, string tag)
    {
        int count=0;
        for(int i=1;i<=(int)shapes.Count;i++)
        {
            dynamic shape=shapes.Item(i);
            if(!string.IsNullOrEmpty(TagValue(shape,tag))) count++;
            if((int)shape.Type==6) count+=CountTagged(shape.GroupItems,tag);
        }
        return count;
    }
    private static int CountNumberLabels(dynamic shapes) { return CountTagged(shapes,"LABPHOTO_NUMBER_STYLE"); }
    private static int CountAttachedPhotos(dynamic shapes) { return CountTagged(shapes,"LABPHOTO_ATTACHED_LABEL"); }
    private static void CheckAttachedPair(dynamic photo, dynamic label, string message)
    {
        Check(TagValue(photo,"LABPHOTO_ATTACHED_LABEL")=="1",message+": photo tag");
        dynamic group=photo.ParentGroup;
        Check((int)group.Type==6 && (int)group.GroupItems.Count==2,message+": photo and label are grouped");
        bool labelFound=false;
        for(int i=1;i<=(int)group.GroupItems.Count;i++) if((int)group.GroupItems.Item(i).Id==(int)label.Id) labelFound=true;
        Check(labelFound,message+": label is in the photo group");
        Check(Math.Abs((double)label.Left-(double)photo.Left)<.05 && Math.Abs((double)label.Top-(double)photo.Top)<.05,message+": label begins at photo top left");
    }
    private static bool SameBounds(dynamic shape, RectangleF expected)
    {
        RectangleF actual=Bounds(shape);
        return Math.Abs(actual.Left-expected.Left)<.05 && Math.Abs(actual.Top-expected.Top)<.05 && Math.Abs(actual.Width-expected.Width)<.05 && Math.Abs(actual.Height-expected.Height)<.05;
    }
    private static void CheckUnchanged(dynamic shape, RectangleF expected, string message) { Check(SameBounds(shape,expected),message); }
    private static void TestPhotoLabelAttachment(dynamic app, dynamic presentation, PowerPointHost host, string imagePath)
    {
        dynamic slide=presentation.Slides.Add((int)presentation.Slides.Count+1,12); app.ActiveWindow.View.GotoSlide((int)slide.SlideIndex);
        // Add in a different order from the visual order to prove row sorting
        // uses the slide position, including an imperfectly level first row.
        dynamic bottomRight=slide.Shapes.AddPicture(imagePath,0,-1,260f,235f,110f,80f);
        dynamic topRight=slide.Shapes.AddPicture(imagePath,0,-1,250f,45f,110f,80f);
        dynamic topLeft=slide.Shapes.AddPicture(imagePath,0,-1,45f,52f,110f,80f);
        dynamic first=host.AddNumberLabel("square",0);
        Check((string)first.TextFrame2.TextRange.Text=="0","Square label text"); CheckAttachedPair(topLeft,first,"First visual photo");
        dynamic second=host.AddNumberLabel("circle",0);
        CheckAttachedPair(topRight,second,"Second visual photo");
        dynamic third=host.AddNumberLabel("paren",0);
        Check((string)third.TextFrame2.TextRange.Text=="(A)","Alphabet label begins at A"); CheckAttachedPair(bottomRight,third,"Third visual photo");
        Console.WriteLine("Actual PowerPoint: labels attach top-left in visual reading order and group with photos passed.");
    }
    private static void TestNumberFormatting(dynamic app, dynamic presentation, PowerPointHost host, string output, string imagePath)
    {
        NumberLabelSettings[] cases = {
            new NumberLabelSettings {FontName="Arial",FontSize=12,TextColorArgb=Color.DarkRed.ToArgb()},
            new NumberLabelSettings {FontName="맑은 고딕",FontSize=28.5f,TextColorArgb=Color.FromArgb(12,98,205).ToArgb()},
            new NumberLabelSettings {FontName="Times New Roman",FontSize=48,TextColorArgb=Color.ForestGreen.ToArgb()}
        };
        int caseIndex=0;
        foreach(NumberLabelSettings settings in cases)
        {
            dynamic slide=presentation.Slides.Add((int)presentation.Slides.Count+1,12);
            app.ActiveWindow.View.GotoSlide((int)slide.SlideIndex);
            for(int i=0;i<6;i++) slide.Shapes.AddPicture(imagePath,0,-1,40f+i*145f,85f,115f,85f);
            foreach(string style in new[]{"square","circle","paren","suffix"})
            {
                dynamic shape=host.AddNumberLabel(style,1234,settings);
                dynamic frame=shape.TextFrame2;
                Check(Math.Abs((double)shape.Width-(double)shape.Height)<.01,"Custom-font labels must remain square");
                Check(Math.Abs((double)frame.TextRange.Font.Size-settings.FontSize)<.01,"Keep exact requested point size");
                Check((string)frame.TextRange.Font.Name==settings.FontName,"Apply selected font");
                Check((int)frame.TextRange.Font.Fill.ForeColor.RGB==settings.OfficeColor,"Apply selected RGB color");
                Check((int)frame.AutoSize==0,"AutoFit disabled");
                Check((double)frame.TextRange.BoundWidth <= (double)shape.Width-2*settings.Padding+.1,"Text fits square horizontally");
                Check((double)frame.TextRange.BoundHeight <= (double)shape.Height-2*settings.Padding+.1,"Text fits square vertically");
                Check((double)frame.TextRange.BoundLeft >= (double)shape.Left+settings.Padding-.1,"Text remains inside left edge");
                Check((double)frame.TextRange.BoundTop >= (double)shape.Top+settings.Padding-.1,"Text remains inside top edge");
                if(style=="paren" || style=="suffix") Check((int)shape.Fill.Visible==0 && (int)shape.Line.Visible==0,"Transparent square text box");
            }
            if(caseIndex==1)
            {
                dynamic next=host.AddNumberLabel("paren",null,settings);
                Check((string)next.TextFrame2.TextRange.Text==NumberLabels.Text("paren",1235) && Math.Abs((double)next.TextFrame2.TextRange.Font.Size-28.5)<.01,"Next-alphabet action also uses specified formatting");
            }
            slide.Export(Path.Combine(output,"number-format-"+caseIndex+".png"),"PNG",1440,810);
            int count=(int)slide.Shapes.Count;
            try
            {
                host.AddNumberLabel("paren",int.MaxValue,new NumberLabelSettings {FontSize=400});
                Check(false,"Oversized label should fail with no partial shape");
            }
            catch(InvalidOperationException) { Check((int)slide.Shapes.Count==count,"Failed insertion removes its provisional shape"); }
            caseIndex++;
        }
        TestProportionalNumberContainers(app,presentation,host,imagePath);
        Console.WriteLine("Actual PowerPoint: square containers, 3 fonts/sizes/colors, fixed font sizes, long labels and failure cleanup passed.");
    }
    private static void ValidatePlan(SelectionSnapshot after)
    {
        List<RectangleF> boxes=new List<RectangleF>();
        Dictionary<int, List<RectangleF>> rows=new Dictionary<int, List<RectangleF>>();
        foreach(PhotoSnapshot p in after.Photos)
        {
            double r=p.Rotation*Math.PI/180;
            double w=Math.Abs(p.Width*Math.Cos(r))+Math.Abs(p.Height*Math.Sin(r));
            double h=Math.Abs(p.Width*Math.Sin(r))+Math.Abs(p.Height*Math.Cos(r));
            double left=p.Left+(p.Width-w)/2, top=p.Top+(p.Height-h)/2;
            Check(left>=95.95 && left+w<=864.05 && top>=53.95 && top+h<=486.05,"Margins including rotation");
            RectangleF box=new RectangleF((float)left,(float)top,(float)w,(float)h);
            foreach(RectangleF other in boxes) Check(!box.IntersectsWith(other),"Actual photos overlap");
            boxes.Add(box);
            int rowKey=(int)Math.Round(top*100);
            if(!rows.ContainsKey(rowKey)) rows.Add(rowKey,new List<RectangleF>());
            rows[rowKey].Add(box);
        }
        double leftEdge=double.PositiveInfinity;
        foreach(RectangleF box in boxes) leftEdge=Math.Min(leftEdge,box.Left);
        int largestRow=0;
        int lastRow=int.MinValue;
        foreach(var entry in rows) { largestRow=Math.Max(largestRow,entry.Value.Count);lastRow=Math.Max(lastRow,entry.Key); }
        foreach(var entry in rows)
        {
            double rowLeft=double.PositiveInfinity;
            foreach(RectangleF box in entry.Value) rowLeft=Math.Min(rowLeft,box.Left);
            Check(Math.Abs(rowLeft-leftEdge)<.05,"Actual rows share left edge");
            Check(entry.Key==lastRow || entry.Value.Count==largestRow,"Only final actual row may have empty slots");
        }
    }
}


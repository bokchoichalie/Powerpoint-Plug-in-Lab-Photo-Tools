using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using System.Xml;
using LabPhotoTools;

internal static class TestSuite
{
    private static int checks;
    private static string output;
    private static void Check(bool condition, string message)
    {
        checks++;
        if (!condition) throw new Exception(message);
    }
    [STAThread]
    private static int Main(string[] args)
    {
        try
        {
            output = args[0]; Directory.CreateDirectory(output);
            TestLayout(); TestRibbon(); TestNumbers();
            Application.EnableVisualStyles();
            TestNumberSettings();
            if (args.Length < 2 || args[1] != "--skip-dialogs") TestDialogs();
            Console.WriteLine("PASS: " + checks + " assertions."); return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }
    private static List<PhotoBox> Photos(int count, int seed)
    {
        Random random = new Random(seed);
        List<PhotoBox> photos = new List<PhotoBox>();
        for (int i = 0; i < count; i++) photos.Add(new PhotoBox { Id = i+1, Left = (i%7)*30, Top = (i/7)*35, Width = 40+random.Next(200), Height = 40+random.Next(200) });
        return photos;
    }
    private static void TestLayout()
    {
        foreach (Size slide in new[] { new Size(960,540), new Size(720,540), new Size(540,960), new Size(600,600) })
        foreach (int count in new[] { 1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,50,101,1000 })
        {
            List<PhotoBox> photos = Photos(count, count);
            List<PhotoBox> plan = LabPhotoTools.Layout.SmartArrange(photos, slide.Width, slide.Height);
            Check(plan.Count == count && plan.Select(p=>p.Id).Distinct().Count() == count, "Photo set changed");
            var rowGroups = plan.GroupBy(p=>p.Top).ToList();
            int columns = rowGroups[0].Count();
            for (int r=0; r<rowGroups.Count; r++)
            {
                Check(Math.Abs(rowGroups[r].First().Left-rowGroups[0].First().Left)<1e-8,"Rows must be left-aligned");
                Check(r==rowGroups.Count-1 ? rowGroups[r].Count()<=columns : rowGroups[r].Count()==columns,"Only final row may be incomplete");
            }
            foreach (PhotoBox p in plan)
            {
                PhotoBox original = photos.Find(x=>x.Id==p.Id);
                Check(Math.Abs(p.Width/p.Height-original.Width/original.Height)<1e-8, "Aspect distorted");
                Check(p.Left >= slide.Width*.1-1e-6 && p.Top >= slide.Height*.1-1e-6 && p.Left+p.Width <= slide.Width*.9+1e-6 && p.Top+p.Height <= slide.Height*.9+1e-6, "20% margins violated");
                Check(original.Width>=40 && original.Height>=40, "Input mutated");
            }
            for (int i=0; i<plan.Count; i++) for(int j=i+1;j<plan.Count;j++)
                Check(!(plan[i].Left < plan[j].Left+plan[j].Width-1e-6 && plan[j].Left < plan[i].Left+plan[i].Width-1e-6 && plan[i].Top < plan[j].Top+plan[j].Height-1e-6 && plan[j].Top < plan[i].Top+plan[i].Height-1e-6), "Photos overlap");
        }
        foreach (KeyValuePair<int,int[]> layout in new Dictionary<int,int[]> {
            {4,new[]{2,2}}, {5,new[]{3,2}}, {6,new[]{3,3}}, {7,new[]{4,3}}, {8,new[]{4,4}},
            {9,new[]{3,3,3}}, {10,new[]{5,5}}, {11,new[]{4,4,3}}, {12,new[]{4,4,4}},
            {13,new[]{4,4,4,1}}, {14,new[]{4,4,4,2}}, {15,new[]{5,5,5}}, {16,new[]{4,4,4,4}}
        })
        {
            List<PhotoBox> squares = Enumerable.Range(0,layout.Key).Select(i=>new PhotoBox {Id=i,Width=100,Height=100,Left=i*10,Top=0}).ToList();
            foreach (Size slide in new[]{new Size(960,540),new Size(720,960)})
            {
                List<PhotoBox> arranged = LabPhotoTools.Layout.SmartArrange(squares,slide.Width,slide.Height);
                Check(arranged.GroupBy(p=>p.Top).Select(row=>row.Count()).SequenceEqual(layout.Value), layout.Key+" photos must use the requested row pattern regardless of slide orientation");
            }
        }
        Console.WriteLine("Layout matrix passed (including 1,000 photos and mixed portrait/landscape).");
    }
    private static void TestRibbon()
    {
        Connect connect = new Connect();
        XmlDocument xml = new XmlDocument(); xml.LoadXml(connect.GetCustomUI("test"));
        XmlNamespaceManager ns = new XmlNamespaceManager(xml.NameTable); ns.AddNamespace("r", xml.DocumentElement.NamespaceURI);
        XmlNodeList groups = xml.SelectNodes("//r:group",ns);
        Check(groups[1].Attributes["label"].Value=="배치" && groups[2].Attributes["label"].Value=="빠른 번호 매기기" && groups[3].Attributes["label"].Value=="치수측정" && groups[4].Attributes["label"].Value=="도움말", "Ribbon order");
        Check(xml.SelectNodes("//r:button[starts-with(@id,'number_')]",ns).Count==60,"Expected 55 presets and 5 next buttons");
        XmlNode numbers=xml.SelectSingleNode("//r:group[@id='labPhotoNumbers']",ns);
        Check(numbers.SelectNodes("r:box[starts-with(@id,'numberBox_')]//r:menu | r:box[starts-with(@id,'numberBox_')]//r:gallery | r:box[starts-with(@id,'numberBox_')]//r:dropDown | .//r:toggleButton",ns).Count==0,"Number presets must remain visible");
        Check(numbers.SelectNodes("r:box[starts-with(@id,'numberBox_')]",ns).Count==5,"Five directly accessible style blocks");
        Check(numbers.SelectNodes("r:box[@id='numberFormatBox']/r:dropDown[@id='labNumberFont']",ns).Count==1,"Font selection must be directly on the ribbon");
        Check(numbers.SelectNodes("r:box[@id='numberFormatBox']/r:editBox[@id='labNumberFontSize']",ns).Count==1,"Size input must be directly on the ribbon");
        Check(numbers.SelectNodes("r:box[@id='numberFormatBox']//r:gallery[@id='labNumberColor']",ns).Count==1,"Color selection must be on the ribbon");
        Check(numbers.SelectNodes(".//*[@onAction='OpenNumberSettings']",ns).Count==0,"No separate settings dialog");
        foreach(string style in new[]{"square","circle","paren","suffix","alphaSuffix"})
        {
            XmlNode box=numbers.SelectSingleNode("r:box[@id='numberBox_"+style+"']",ns);
            Check(box.SelectNodes("r:buttonGroup",ns).Count==3,"Three visible rows per number style");
            for(int n=0;n<=10;n++)
            {
                XmlNode button=box.SelectSingleNode("r:buttonGroup/r:button[@tag='"+style+":"+n+"']",ns);
                Check(button!=null && button.Attributes["label"].Value==NumberLabels.Caption(style,n),"Visible numeric preset with correct tag and label");
                Check(button.Attributes["showLabel"]==null || button.Attributes["showLabel"].Value!="false","Number label must be shown");
            }
            XmlNode next=box.SelectSingleNode("r:buttonGroup/r:button[@tag='"+style+":next']",ns);
            Check(next!=null && next.Attributes["label"].Value==(NumberLabels.IsAlphabet(style) ? "L+" : "11+"),"One correctly labelled next button per style");
        }
        HashSet<string> ids=new HashSet<string>();
        foreach(XmlNode node in xml.SelectNodes("//*[@id]")) Check(ids.Add(node.Attributes["id"].Value),"Duplicate id");
        foreach(XmlNode node in xml.SelectNodes("//*[@onAction]")) Check(typeof(Connect).GetMethod(node.Attributes["onAction"].Value)!=null,"Missing callback");
        foreach(XmlNode node in xml.SelectNodes("//*")) foreach(XmlAttribute attribute in node.Attributes)
            if(attribute.Name.StartsWith("get",StringComparison.Ordinal) || attribute.Name=="onChange" || attribute.Name=="onLoad")
                Check(typeof(Connect).GetMethod(attribute.Value)!=null,"Missing settings callback "+attribute.Value);
        string[] icons={"labPhotoEdit","labPhotoBackground","labPhotoGrid","labPhotoSpacing","labPhotoMagic","number_square","number_circle","number_paren","number_suffix"};
        using(Bitmap gallery=new Bitmap(icons.Length*100,110))
        using(Graphics g=Graphics.FromImage(gallery))
        {
            g.Clear(Color.White);
            for(int i=0;i<icons.Length;i++) using(Bitmap icon=RibbonIcons.Draw(icons[i]))
            {
                g.DrawImage(icon,i*100+18,10); RibbonIcons.Get(icons[i]);
                g.DrawString(icons[i].Replace("labPhoto","").Replace("number_",""),SystemFonts.DefaultFont,Brushes.Black,i*100+5,82);
            }
            gallery.Save(Path.Combine(output,"icons.png"),ImageFormat.Png);
        }
        File.WriteAllText(Path.Combine(output,"ribbon.xml"),xml.OuterXml);
        Console.WriteLine("Ribbon XML, callbacks and native icon conversion passed.");
    }
    private static void TestNumbers()
    {
        Check(NumberLabels.Next(new int[0])==11,"First next");
        Check(NumberLabels.Next(new[]{0,10,11,14})==15,"Next sequence");
        Check(NumberLabels.Text("paren",0)=="(A)" && NumberLabels.Text("paren",10)=="(K)" && NumberLabels.Text("paren",11)=="(L)","Alphabet presets");
        Check(NumberLabels.Text("paren",26)=="(AA)","Alphabet continues after Z");
        Check(NumberLabels.Text("suffix",12)=="12)","Suffix");
        Check(NumberLabels.Text("alphaSuffix",0)=="A)"&&NumberLabels.Text("alphaSuffix",10)=="K)"&&NumberLabels.Text("alphaSuffix",11)=="L)"&&NumberLabels.Text("alphaSuffix",26)=="AA)","Alphabet suffix presets and continuation");
    }
    private static void TestNumberSettings()
    {
        string folder=Path.Combine(output,"settings-test"); Directory.CreateDirectory(folder);
        string path=Path.Combine(folder,"preferences.json");
        if(File.Exists(path)) File.Delete(path);
        Check(NumberLabelPreferences.Load(path).FontSize==18,"Default label font size");
        NumberLabelSettings settings=new NumberLabelSettings {FontName="맑은 고딕",FontSize=31.5f,TextColorArgb=Color.FromArgb(12,98,205).ToArgb()};
        NumberLabelPreferences.Save(path,settings);
        NumberLabelSettings saved=NumberLabelPreferences.Load(path);
        Check(saved.FontName==settings.FontName && saved.FontSize==31.5f && saved.TextColorArgb==settings.TextColorArgb,"Font preferences round trip");
        Check(saved.OfficeColor==(12|(98<<8)|(205<<16)),"Office RGB byte order");
        Check(saved.BorderMode=="default"&&saved.FillMode=="default"&&saved.ShowFill(true)&&!saved.ShowFill(false),"Legacy container appearance");
        settings.BorderMode="color";settings.BorderColorArgb=Color.Red.ToArgb();settings.BorderWidth=2.5f;settings.FillMode="none";
        NumberLabelPreferences.Save(path,settings);saved=NumberLabelPreferences.Load(path);
        Check(saved.BorderWidth==2.5f&&saved.BorderColorArgb==Color.Red.ToArgb()&&saved.ShowBorder(false)&&!saved.ShowFill(true),"Border and transparent fill round trip");
        File.WriteAllText(path,"{\"FontName\":\"Arial\",\"FontSize\":17,\"TextColorArgb\":-65536}");saved=NumberLabelPreferences.Load(path);
        Check(saved.FontSize==17&&saved.TextColorArgb==Color.Red.ToArgb()&&saved.BorderWidth==1&&saved.FillMode=="default","Existing three-field preferences migrate without resetting font");
        settings.FontSize=24; NumberLabelPreferences.Save(path,settings);
        Check(NumberLabelPreferences.Load(path).FontSize==24,"Atomic preference replacement");
        File.WriteAllText(path,"broken json"); Check(NumberLabelPreferences.Load(path).FontSize==18,"Damaged preferences fall back safely");
        File.WriteAllText(path,"{\"FontName\":\"Arial\",\"FontSize\":0}"); Check(NumberLabelPreferences.Load(path).FontSize==18,"Invalid size falls back safely");
        NumberLabelSettings seventeen=new NumberLabelSettings {FontSize=17};
        NumberLabelSettings thirtyFour=new NumberLabelSettings {FontSize=34};
        double expected17=17*.5*72/25.4;
        Check(Math.Abs(seventeen.BaseContainerSide-expected17)<.001,"17 pt must use an 0.85 cm container");
        Check(Math.Abs(thirtyFour.BaseContainerSide-expected17*2)<.001,"Container side must stay proportional to point size");
        Check(Math.Abs(seventeen.SquareSide(0,0)-expected17)<.001,"Base container must not have a fixed minimum");
        Check(settings.SquareSide(150,30)>=150+2*settings.Padding,"Long labels grow without shrinking font");
        TestNumberRibbonSettings();
        Console.WriteLine("Number preferences and inline ribbon formatting passed.");
    }
    private static void TestNumberRibbonSettings()
    {
        NumberLabelSettings saved=new NumberLabelSettings();
        Connect connect=new Connect(()=>saved.Copy(),s=>saved=s.Copy());
        Check(connect.GetNumberFontCount(null)>0,"Installed fonts missing");
        int fontIndex=Array.IndexOf(NumberRibbonSettings.Fonts,"Arial");
        Check(fontIndex>=0,"Arial missing");
        connect.SetNumberFont(null,"",fontIndex);
        Check(connect.GetNumberFontLabel(null,connect.GetNumberFontIndex(null))==saved.FontName,"Font selection did not persist");
        connect.SetNumberFontSize(null,"28.5");
        Check(saved.FontSize==28.5f && saved.FontName=="Arial","Size callback changed unrelated settings");
        connect.SetNumberColor(null,"",9);
        Check(saved.TextColor.ToArgb()==Color.Red.ToArgb() && saved.FontSize==28.5f,"Palette callback did not preserve size");
        Check(connect.GetNumberColorHex(null)=="FF0000","Palette did not update displayed color code");
        connect.SetNumberColorHex(null,"#0c62cd");
        Check(saved.OfficeColor==(12|(98<<8)|(205<<16)) && connect.GetNumberColorHex(null)=="0C62CD","Custom color code not applied");
        Check(connect.GetNumberColorImage(null)!=null,"Current color swatch missing");
        dynamic border=new System.Dynamic.ExpandoObject();border.Id="labNumberBorderColor";
        dynamic fill=new System.Dynamic.ExpandoObject();fill.Id="labNumberFillColor";
        connect.SetNumberBoxColorHex(border,"#cc2211");connect.SetNumberBorderWidth(null,"2.75");connect.SetNumberBoxColor(fill,"",3);
        Check(saved.BorderWidth==2.75f&&saved.BorderMode=="color"&&saved.FillMode=="color"&&saved.FillColorArgb==Color.White.ToArgb(),"Container ribbon callbacks");
        connect.SetNumberBoxColor(fill,"",1);Check(!saved.ShowFill(true)&&connect.GetNumberBoxColorHex(fill)=="없음","Transparent fill choice");
        connect.SetNumberBorderWidth(null,"0");Check(!saved.ShowBorder(true),"Zero border width disables border");
        connect.SetNumberBoxColor(border,"",11);Check(saved.BorderWidth==1&&saved.BorderColorArgb==Color.Red.ToArgb(),"Re-enable border through palette");
        Check(connect.GetNumberBoxColorImage(fill)!=null&&connect.GetNumberBoxColorCount(fill)==34,"Box palette default/none/color entries");
        Check(connect.GetNumberColorCount(null)==32,"Expected 32 directly selectable colors");
        for(int i=0;i<connect.GetNumberColorCount(null);i++)
        {
            Check(connect.GetNumberColorItemImage(null,i)!=null,"Palette swatch missing");
            Check(connect.GetNumberColorLabel(null,i).Contains(NumberRibbonSettings.Hex(NumberRibbonSettings.Colors[i])),"Palette name and code disagree");
        }
        foreach(string text in new[]{"", "0", "401", "NaN", "Infinity", "abc"})
        {
            bool failed=false; try{NumberRibbonSettings.ParseSize(text);}catch(ArgumentException){failed=true;}
            Check(failed,"Invalid size accepted: "+text);
        }
        foreach(string text in new[]{"", "GGGGGG", "00000", "0000000", "-00001"})
        {
            bool failed=false; try{NumberRibbonSettings.ParseColor(text);}catch(ArgumentException){failed=true;}
            Check(failed,"Invalid color accepted: "+text);
        }
        System.Globalization.CultureInfo previous=System.Threading.Thread.CurrentThread.CurrentCulture;
        try
        {
            System.Threading.Thread.CurrentThread.CurrentCulture=System.Globalization.CultureInfo.GetCultureInfo("fr-FR");
            Check(NumberRibbonSettings.ParseSize("28,5")==28.5f && NumberRibbonSettings.ParseSize("28.5")==28.5f,"Decimal sizes must support locale and invariant notation");
        }
        finally {System.Threading.Thread.CurrentThread.CurrentCulture=previous;}
        string callbackPath=Path.Combine(output,"settings-test","ribbon-preferences.json");
        NumberLabelPreferences.Save(callbackPath,saved);
        Connect reopened=new Connect(()=>NumberLabelPreferences.Load(callbackPath),s=>NumberLabelPreferences.Save(callbackPath,s));
        Check(reopened.GetNumberFontSize(null)==connect.GetNumberFontSize(null) && reopened.GetNumberColorHex(null)=="0C62CD","Reopened ribbon must show persisted formatting");
    }
    private sealed class PreviewHost : PowerPointHost
    {
        public PreviewHost() : base(new object()) { }
        internal override ImagePreviewSession CreatePreviewSession(SelectionSnapshot selection)
        {
            string job = Engine.NewJobDirectory();
            using(Bitmap bitmap=new Bitmap(720,480))
            using(Graphics g=Graphics.FromImage(bitmap))
            {
                g.Clear(Color.FromArgb(208,225,234));
                g.FillEllipse(Brushes.DarkSlateBlue,220,100,210,210);
                g.FillRectangle(Brushes.SteelBlue,90,320,530,35);
                bitmap.Save(Path.Combine(job,"input.png"));
            }
            ImagePreviewSession session = new ImagePreviewSession(job);
            typeof(ImagePreviewSession).GetProperty("BackgroundRemoved").SetValue(session,new Bitmap(session.Original),null);
            typeof(ImagePreviewSession).GetField("backgroundTask",BindingFlags.Instance|BindingFlags.NonPublic).SetValue(session,System.Threading.Tasks.Task.FromResult(0));
            return session;
        }
    }
    private static List<Control> Tree(Control root)
    {
        List<Control> controls=new List<Control>();
        foreach(Control c in root.Controls) { controls.Add(c); controls.AddRange(Tree(c)); }
        return controls;
    }
    private static void TestDialogs()
    {
        SelectionSnapshot snapshot = new SelectionSnapshot { SlideWidth=960,SlideHeight=540 };
        snapshot.Photos.Add(new PhotoSnapshot { Id=1,Width=180,Height=120 });
        int cases=0;
        foreach(string mode in new[]{"image","layout","spacing","background"})
        foreach(float scale in new[]{1f,1.25f,1.5f,2f,2.5f,3f})
        foreach(Size size in new[]{new Size(800,600),new Size(1024,768),new Size(1366,768),new Size(1920,1080),new Size(3840,2160)})
        {
            using(ToolForm form=new ToolForm(new PreviewHost(),snapshot,mode))
            {
                form.Opacity=0; form.ShowInTaskbar=false;
                // Layout without Shown/engine side effects. Explicitly scale fonts
                // and metrics to model 100-300% display/text scaling.
                form.Show();
                Application.DoEvents();
                List<Control> controls=Tree(form);
                List<Font> fonts=controls.Select(c=>new Font(c.Font.FontFamily,c.Font.SizeInPoints*scale,c.Font.Style)).ToList();
                form.AutoScaleMode=AutoScaleMode.None;
                form.Scale(new SizeF(scale,scale));
                form.Font=new Font("맑은 고딕",10*scale);
                for(int i=0;i<controls.Count;i++) controls[i].Font=fonts[i];
                form.MinimumSize=Size.Empty;
                form.ClientSize=new Size(size.Width-32,size.Height-90);
                form.PerformLayout();
                foreach(Control c in controls) c.PerformLayout();
                form.PerformLayout();
                string context=mode+" "+size+" "+scale;
                foreach(TableLayoutPanel panel in controls.OfType<TableLayoutPanel>())
                {
                    List<Control> children=panel.Controls.Cast<Control>().ToList();
                    for(int i=0;i<children.Count;i++) for(int j=i+1;j<children.Count;j++)
                        Check(!children[i].Bounds.IntersectsWith(children[j].Bounds),"Controls overlap: "+context+" "+children[i].Text+" / "+children[j].Text);
                }
                foreach(NumericUpDown number in controls.OfType<NumericUpDown>())
                    Check(number.Width>60 && number.Height>=number.PreferredHeight,"Input clipped: "+context);
                foreach(Label label in controls.OfType<Label>())
                    Check(label.Height >= label.GetPreferredSize(new Size(label.Width,0)).Height, "Text clipped: "+context+" "+label.Text);
                Panel viewport=controls.OfType<Panel>().First(c=>c.GetType()==typeof(Panel) && c.AutoScroll);
                Check(viewport.ClientSize.Height>=30,"No space to scroll content: "+context);
                foreach(Button button in controls.OfType<Button>().Where(b=>b.Parent is FlowLayoutPanel))
                {
                    Rectangle bounds=form.RectangleToClient(button.RectangleToScreen(button.ClientRectangle));
                    Check(form.ClientRectangle.Contains(bounds),"Footer button outside form: "+context);
                }
                Check(!controls.OfType<Label>().Any(c=>c.Text=="배경 제거"),"Background should be separate");
                if(mode=="image" && ((scale==1 && size.Width==1366)||(scale==2 && size.Width==1024)))
                {
                    PhotoPreview preview=controls.OfType<PhotoPreview>().Single();
                    using(Bitmap source=new Bitmap(400,260))
                    {
                        using(Graphics g=Graphics.FromImage(source)) { g.Clear(Color.LightSteelBlue);g.FillEllipse(Brushes.SteelBlue,110,50,170,170); }
                        preview.Source=source;
                        using(Bitmap bitmap=new Bitmap(form.Width,form.Height)) { form.DrawToBitmap(bitmap,new Rectangle(Point.Empty,bitmap.Size));bitmap.Save(Path.Combine(output,"dialog-"+size.Width+"-"+scale+".png"),ImageFormat.Png); }
                        preview.Source=null;
                    }
                }
                cases++;
            }
        }
        Console.WriteLine("Dialog layout matrix passed: "+cases+" cases, 800x600 to 4K, 100-300%.");
    }
}

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace LabPhotoTools
{
    public sealed class MeasurementForm : Form
    {
        private readonly PowerPointHost host;
        private readonly SelectionSnapshot selection;
        private readonly MeasurementSession session;
        internal readonly MeasurementCanvas Canvas=new MeasurementCanvas();
        private readonly MeasurementMagnifier magnifier=new MeasurementMagnifier();
        private readonly TableLayoutPanel sidebar=new TableLayoutPanel {Dock=DockStyle.Fill,Margin=Padding.Empty,ColumnCount=1,RowCount=2};
        private readonly Panel settingsViewport=new Panel {Dock=DockStyle.Fill,AutoScroll=true,Margin=Padding.Empty};
        private readonly TableLayoutPanel settings=new TableLayoutPanel {ColumnCount=1,Margin=Padding.Empty,Padding=new Padding(8)};
        private readonly Label status=new Label {AutoSize=true,Dock=DockStyle.Fill,Margin=new Padding(6),Text="사진의 스케일바 양 끝을 지정해 주세요."};
        private readonly Label scaleStatus=new Label {AutoSize=true,Dock=DockStyle.Top,Margin=new Padding(4)};
        private readonly NumericUpDown actual=new NumericUpDown {DecimalPlaces=6,Minimum=.000001m,Maximum=1000000000,Value=100,Width=140};
        private readonly ComboBox units=new ComboBox {DropDownStyle=ComboBoxStyle.DropDownList,Width=65};
        private readonly NumericUpDown lineWidth=new NumericUpDown {Minimum=1,Maximum=12,Value=2,Width=60};
        private readonly NumericUpDown fontSize=new NumericUpDown {Minimum=6,Maximum=72,Value=13,Width=60};
        private readonly CheckBox dashed=new CheckBox {Text="점선",AutoSize=true}, filled=new CheckBox {Text="채우기",AutoSize=true}, guide=new CheckBox {Text="보조 도형 (값 숨김)",AutoSize=true}, obtuse=new CheckBox {Text="수평·수직 각도: 둔각",AutoSize=true};
        private readonly TextBox annotation=new TextBox {Width=260,Text="메모",AccessibleName="주석 텍스트"};
        private readonly ListView results=new ListView {View=View.Details,FullRowSelect=true,MultiSelect=true,HideSelection=false,Dock=DockStyle.Top,Height=185};
        private readonly List<Button> toolButtons=new List<Button>();
        private readonly ToolTip tips=new ToolTip();
        private MeasurementItem pendingCalibration;
        private bool reflowing,refreshing;
        private bool footerInScroll;
        private readonly Button apply;
        private readonly List<Image> buttonImages=new List<Image>();
        private Label navigationHint;
        private FlowLayoutPanel footer;
        private FlowLayoutPanel calibrationTools;
        private readonly Panel windowViewport=new Panel {Dock=DockStyle.Fill};
        private TableLayoutPanel root;
        private readonly TableLayoutPanel calibrationBand=new TableLayoutPanel {Dock=DockStyle.Top,ColumnCount=2,RowCount=1,Margin=Padding.Empty};
        private readonly TableLayoutPanel calibrationValues=new TableLayoutPanel {Dock=DockStyle.Top,ColumnCount=1,Margin=Padding.Empty};
        public MeasurementForm(PowerPointHost host,SelectionSnapshot selection) : this(host,selection,host.CreateMeasurementSession(selection)) { }
        internal MeasurementForm(PowerPointHost host,SelectionSnapshot selection,MeasurementSession session)
        {
            this.host=host;this.selection=selection;this.session=session;
            Text="Lab Photo Tools · 치수측정";Font=new Font("맑은 고딕",9.5f);AutoScaleDimensions=new SizeF(96,96);AutoScaleMode=AutoScaleMode.Dpi;
            StartPosition=FormStartPosition.CenterParent;ClientSize=new Size(1280,830);MinimumSize=new Size(380,340);BackColor=Color.FromArgb(248,249,251);MinimizeBox=false;MaximizeBox=false;FormBorderStyle=FormBorderStyle.FixedDialog;
            root=new TableLayoutPanel {Dock=DockStyle.Fill,ColumnCount=2,RowCount=1,Padding=new Padding(8),Margin=Padding.Empty};
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,500));root.RowStyles.Add(new RowStyle(SizeType.Percent,100));Controls.Add(windowViewport);windowViewport.Controls.Add(root);
            Canvas.Margin=new Padding(0,0,8,0);root.Controls.Add(Canvas,0,0);root.Controls.Add(sidebar,1,0);
            sidebar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));sidebar.RowStyles.Add(new RowStyle(SizeType.Percent,100));sidebar.RowStyles.Add(new RowStyle(SizeType.AutoSize));sidebar.Controls.Add(settingsViewport,0,0);
            settingsViewport.Controls.Add(settings);settings.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));
            FlowLayoutPanel top=Flow();top.Controls.Add(Label("치수측정",true));
            top.Controls.Add(Button("전체 보기",delegate{Canvas.ResetView();Canvas.Focus();}));top.Controls.Add(Button("확대 +",delegate{Canvas.Zoom(1.25);Canvas.Focus();}));top.Controls.Add(Button("축소 −",delegate{Canvas.Zoom(.8);Canvas.Focus();}));Add(top);
            navigationHint=Label("휠: 확대 · 가운데 버튼: 이동\nWASD/방향키: 1 px · Shift: 10 px · Enter: 점 확정",false);navigationHint.Dock=DockStyle.Top;Add(navigationHint);
            Canvas.Document=session.Document;Canvas.Source=session.Image;Canvas.Rotation=session.Rotation;magnifier.Source=session.Image;magnifier.Document=session.Document;magnifier.Rotation=session.Rotation;magnifier.Point=new MeasurePoint(session.Document.Width/2,session.Document.Height/2);
            BuildSettings();
            if(Canvas.Document.CalibrationLengthMicrons>0)actual.Value=Math.Max(actual.Minimum,Math.Min(actual.Maximum,(decimal)(Canvas.Document.CalibrationLengthMicrons/MeasurementGeometry.UnitFactor(Canvas.Document.Unit))));
            Add(status);footer=Flow();footer.FlowDirection=FlowDirection.RightToLeft;
            apply=Button("측정 사진 복사",Apply);apply.BackColor=Color.FromArgb(25,98,180);apply.ForeColor=Color.White;
            footer.Controls.Add(apply);footer.Controls.Add(Button("닫기",delegate{Close();}));footer.Controls.Add(Button("CSV 저장",Export));sidebar.Controls.Add(footer,0,1);
            Canvas.CalibrationReady+=delegate(MeasurementItem item){pendingCalibration=item;scaleStatus.Text="기준 "+MeasurementGeometry.ReferenceLength(item).ToString("0.###")+" px · 실제 길이를 입력하고 스케일 적용";status.Text="선택한 기준의 실제 길이를 입력하고 ‘스케일 적용’을 누르세요.";};
            Canvas.Changed+=RefreshResults;Canvas.Status+=delegate(string text){status.Text=text;};
            Canvas.HoverChanged+=delegate(MeasurePoint p){magnifier.Point=p;magnifier.Invalidate();status.Text="커서 X "+(p.X*session.Image.Width/Canvas.Document.Width).ToString("0.##")+" / Y "+(p.Y*session.Image.Height/Canvas.Document.Height).ToString("0.##")+" px · WASD/방향키: 1 px · Shift: 10 px · Enter/클릭: 점 확정";};
            Shown+=delegate{FitScreen();Reflow();Canvas.Focus();};
            windowViewport.SizeChanged+=delegate{Reflow();};
            ClientSizeChanged+=delegate{Reflow();};
            RefreshResults();Canvas.SetTool(Canvas.Document.HasScale?"select":"line",!Canvas.Document.HasScale);
            if(Canvas.Document.HasScale)status.Text="저장된 스케일과 측정 도형을 불러왔습니다.";
            tips.SetToolTip(actual,"스케일바의 실제 길이. 평행선은 간격, 3점원은 지름입니다.");
        }
        private static FlowLayoutPanel Flow(){return new FlowLayoutPanel {AutoSize=true,AutoSizeMode=AutoSizeMode.GrowAndShrink,Dock=DockStyle.Top,WrapContents=true,Margin=Padding.Empty};}
        private Label Label(string text,bool bold){return new Label {Text=text,AutoSize=true,Margin=new Padding(4,8,4,5),Font=bold?new Font(Font,FontStyle.Bold):Font};}
        private Button Button(string text,Action action)
        {Button b=new Button {Text=text,AutoSize=true,AutoSizeMode=AutoSizeMode.GrowAndShrink,Padding=new Padding(6,5,6,5),Margin=new Padding(3),FlatStyle=FlatStyle.Flat,BackColor=Color.White};b.Click+=delegate{try{action();}catch(Exception ex){status.Text=ex.Message;}};return b;}
        private void Add(Control control){int row=settings.Controls.Count;settings.RowStyles.Add(new RowStyle(SizeType.AutoSize));settings.Controls.Add(control,0,row);settings.RowCount=row+1;}
        private void AddCalibration(Control control){int row=calibrationValues.Controls.Count;calibrationValues.RowStyles.Add(new RowStyle(SizeType.AutoSize));calibrationValues.Controls.Add(control,0,row);calibrationValues.RowCount=row+1;}
        private void BuildSettings()
        {
            calibrationBand.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));calibrationBand.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,168));calibrationBand.RowStyles.Add(new RowStyle(SizeType.Percent,100));Add(calibrationBand);
            calibrationValues.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));calibrationBand.Controls.Add(calibrationValues,0,0);
            magnifier.Name="MeasurementMagnifier";magnifier.Size=new Size(160,120);magnifier.Margin=new Padding(4,8,4,8);magnifier.Anchor=AnchorStyles.Top|AnchorStyles.Left;calibrationBand.Controls.Add(magnifier,1,0);
            AddCalibration(Label("1. 수동 스케일",true));calibrationTools=Flow();
            foreach(string kind in new[]{"line","gap","circle3"})
            {string k=kind;Button b=Button(k=="line"?"선":k=="gap"?"평행선":"3점원",delegate{ChooseTool(k,true);});SetImage(b,k);calibrationTools.Controls.Add(b);}
            AddCalibration(calibrationTools);AddCalibration(Label("실제 길이",false));FlowLayoutPanel scaleRow=Flow();actual.AccessibleName="스케일 실제 길이";units.AccessibleName="스케일 단위";scaleRow.Controls.Add(actual);units.Items.AddRange(new object[]{"nm","µm","mm","cm"});units.SelectedItem=Canvas.Document.Unit;scaleRow.Controls.Add(units);AddCalibration(scaleRow);
            FlowLayoutPanel scaleActions=Flow();scaleActions.Controls.Add(Button("스케일 적용",ApplyScale));scaleActions.Controls.Add(Button("표시 단위 변경",delegate{Canvas.PushUndo();Canvas.Document.Unit=(string)units.SelectedItem;RefreshResults();}));AddCalibration(scaleActions);AddCalibration(scaleStatus);
            Add(Label("2. 측정 도구",true));FlowLayoutPanel tools=Flow();
            foreach(KeyValuePair<string,string> pair in MeasurementGeometry.Names)
            {
                string k=pair.Key;Button b=Button(pair.Value,delegate{ChooseTool(k,false);});b.Tag=k;SetImage(b,k);tools.Controls.Add(b);toolButtons.Add(b);
                tips.SetToolTip(b,Hint(k));
            }
            Add(tools);FlowLayoutPanel edit=Flow();
            edit.Controls.Add(Button("선택·이동",delegate{ChooseTool("select",false);}));edit.Controls.Add(Button("전체 선택",Canvas.SelectAll));
            edit.Controls.Add(Button("지우개",delegate{ChooseTool("erase",false);}));edit.Controls.Add(Button("선택 삭제",Canvas.DeleteSelected));
            edit.Controls.Add(Button("전체 삭제",Canvas.ClearAll));edit.Controls.Add(Button("실행 취소",Canvas.Undo));edit.Controls.Add(Button("다시 실행",Canvas.Redo));
            edit.Controls.Add(Button("그리기 완료",Canvas.Finish));Add(edit);
            Add(Label("3. 선·측정값 서식",true));FlowLayoutPanel format=Flow();
            format.Controls.Add(Label("선 굵기",false));format.Controls.Add(lineWidth);format.Controls.Add(Label("글자 크기",false));format.Controls.Add(fontSize);
            format.Controls.Add(Button("선 색",delegate{ChooseColor(false);}));format.Controls.Add(Button("글자 색",delegate{ChooseColor(true);}));
            format.Controls.Add(dashed);format.Controls.Add(filled);format.Controls.Add(guide);format.Controls.Add(obtuse);Add(format);
            Add(annotation);Add(Button("선택 도형에 서식 적용",delegate{SyncStyle();Canvas.ApplyStyle();}));
            foreach(Control c in new Control[]{lineWidth,fontSize})((NumericUpDown)c).ValueChanged+=delegate{SyncStyle();};
            dashed.CheckedChanged+=delegate{SyncStyle();};filled.CheckedChanged+=delegate{SyncStyle();};guide.CheckedChanged+=delegate{SyncStyle();};obtuse.CheckedChanged+=delegate{SyncStyle();};annotation.TextChanged+=delegate{SyncStyle();};
            Add(Label("4. 측정 결과",true));results.Columns.Add("#",34);results.Columns.Add("도구",90);results.Columns.Add("측정값",280);Add(results);
            results.SelectedIndexChanged+=delegate
            {if(refreshing)return;Canvas.Selected.Clear();foreach(ListViewItem row in results.SelectedItems)Canvas.Selected.Add((int)row.Tag);Canvas.Invalidate();};
        }
        private void SyncStyle(){Canvas.Style.LineWidth=(float)lineWidth.Value;Canvas.Style.FontSize=(float)fontSize.Value;Canvas.Style.Dashed=dashed.Checked;Canvas.Style.Filled=filled.Checked;Canvas.Style.Guide=guide.Checked;Canvas.Style.Obtuse=obtuse.Checked;Canvas.Style.Note=annotation.Text;}
        private void ChooseColor(bool text)
        {using(ColorDialog dialog=new ColorDialog {Color=Color.FromArgb(text?Canvas.Style.TextColorArgb:Canvas.Style.ColorArgb),FullOpen=true})if(dialog.ShowDialog(this)==DialogResult.OK){if(text)Canvas.Style.TextColorArgb=dialog.Color.ToArgb();else Canvas.Style.ColorArgb=dialog.Color.ToArgb();}}
        private void ChooseTool(string k,bool calibration)
        {
            if(!calibration&&!Canvas.Document.HasScale&&k!="select"&&k!="erase"){status.Text="먼저 기준을 지정하고 스케일을 적용하세요.";return;}
            SyncStyle();Canvas.SetTool(k,calibration);status.Text=(calibration?"스케일 기준: ":"")+Hint(k);RefreshButtonStates();
        }
        private static string Hint(string k)
        {
            switch(k)
            {
                case "line":return "두 끝점을 클릭하세요.";
                case "gap":case "pointline":return "기준선 2점 → 측정할 위치를 차례로 클릭 → Enter 또는 우클릭으로 완료.";
                case "circle3":return "원 둘레 위의 서로 떨어진 세 점을 클릭하세요.";
                case "circle":return "중심점 → 원 둘레의 한 점을 클릭하세요.";
                case "circle_distance":return "기존 원 두 개의 둘레나 중심을 클릭하세요. Min은 외부 간격, Center는 중심 거리입니다.";
                case "rect":return "사각형의 대각선 두 꼭짓점을 클릭하세요.";
                case "rect3":return "첫 변의 2점 → 높이 방향 1점을 클릭하세요.";
                case "ellipse":return "한 축의 양 끝점 → 다른 축 방향의 1점을 클릭하세요.";
                case "angle":return "첫 점 → 꼭짓점 → 끝점을 클릭하세요.";
                case "hangle":case "vangle":return "시작점 → 끝점을 클릭하세요. ‘둔각’으로 보각을 선택할 수 있습니다.";
                case "curve":return "시작점 → 제어점 → 끝점을 반복해서 클릭 → Enter 또는 우클릭으로 완료.";
                case "polygon":case "polyline":return "점을 차례로 클릭 → Enter 또는 우클릭으로 완료.";
                case "lasso":case "draw":return "마우스 왼쪽 버튼을 누른 채 드래그하고 놓으면 완료됩니다.";
                case "select":return "도형을 끌어 이동 · 측정값을 끌어 위치 조정 · Ctrl+클릭: 여러 개 선택.";
                case "erase":return "삭제할 측정 도형을 클릭하세요. Ctrl+Z로 되돌릴 수 있습니다.";
                case "text":return "주석 입력란의 텍스트를 넣을 위치를 클릭하세요.";
                default:return "사진에서 위치를 클릭하세요. Esc로 그리기를 취소합니다.";
            }
        }
        private void ApplyScale()
        {
            MeasurementItem reference=pendingCalibration??Canvas.Document.Calibration;
            // Validate before saving an undo entry, and preserve old scale on error.
            double pixels=MeasurementGeometry.ReferenceLength(reference);if(pixels<.01)throw new InvalidOperationException("기준을 다시 지정하세요.");
            Canvas.PushUndo();Canvas.Document.Calibrate(reference,(double)actual.Value,(string)units.SelectedItem);pendingCalibration=null;
            status.Text="스케일 적용 완료. 측정 도구를 선택하세요.";Canvas.SetTool("select",false);RefreshResults();
        }
        private void RefreshButtonStates()
        {foreach(Button b in toolButtons){b.Enabled=Canvas.Document.HasScale;b.BackColor=(string)b.Tag==Canvas.Tool&&!Canvas.Calibrating?Color.FromArgb(205,226,249):Color.White;}}
        private void RefreshResults()
        {
            refreshing=true;
            try
            {
                magnifier.Document=Canvas.Document;RefreshButtonStates();
                scaleStatus.Text=pendingCalibration!=null?"기준 "+MeasurementGeometry.ReferenceLength(pendingCalibration).ToString("0.###")+" px · 실제 길이 입력 후 스케일 적용":Canvas.Document.HasScale?"Scale: "+Canvas.Document.MicronsPerPixel.ToString("G6")+" µm/px · 표시 "+Canvas.Document.Unit:"Scale: 미설정 · 기준을 지정하세요.";
                results.BeginUpdate();results.Items.Clear();
                foreach(MeasurementItem i in Canvas.Document.Items)
                {
                    ListViewItem row=new ListViewItem(new[]{i.Id.ToString(),MeasurementGeometry.Names[i.Kind],MeasurementGeometry.Build(i,Canvas.Document).Text.Replace("\n"," / ")}){Tag=i.Id};
                    results.Items.Add(row);row.Selected=Canvas.Selected.Contains(i.Id);
                }
                results.EndUpdate();Canvas.Invalidate();if(apply!=null)apply.Enabled=Canvas.Document.HasScale;
            }
            finally{refreshing=false;}
        }
        private void SetImage(Button button,string kind)
        {Bitmap image=MeasurementIcons.Draw(kind,24);buttonImages.Add(image);button.Image=image;button.TextImageRelation=TextImageRelation.ImageBeforeText;}
        private void Reflow()
        {
            if(reflowing)return;reflowing=true;
            try
            {
                float scale;
                using(Graphics g=CreateGraphics())using(Font baseline=new Font("맑은 고딕",9.5f))
                    scale=Math.Max(DeviceDpi/96f,Font.GetHeight(g)/baseline.GetHeight(96));
                actual.Width=TextRenderer.MeasureText("1000000000.000000",actual.Font).Width+(int)(32*scale);units.Width=Math.Max((int)(78*scale),TextRenderer.MeasureText("mm",units.Font).Width+(int)(36*scale));
                int loupeWidth=Math.Max(4,(int)(160*scale)/4*4);magnifier.Size=new Size(loupeWidth,loupeWidth*3/4);
                int valuesWidth=Math.Max(actual.Width+actual.Margin.Horizontal+units.Width+units.Margin.Horizontal,
                    calibrationTools.Controls.Cast<Control>().Sum(c=>c.PreferredSize.Width+c.Margin.Horizontal));
                valuesWidth+=8;int minimumWidth=valuesWidth+loupeWidth+magnifier.Margin.Horizontal+settings.Padding.Horizontal;
                // Only the sidebar scrolls, even on a small display. The image
                // keeps the entire available height and never moves below panels.
                int available=Math.Max(1,root.ClientSize.Width-root.Padding.Horizontal);
                int desired=Math.Max(minimumWidth+SystemInformation.VerticalScrollBarWidth,(int)Math.Min(available*.30,600*scale));
                int right=Math.Min(desired,(int)(available*.48));root.ColumnStyles[1].Width=Math.Max(1,right);
                foreach(Button b in footer.Controls.OfType<Button>())b.MaximumSize=new Size(Math.Max(60,right-b.Margin.Horizontal),0);
                bool scrollFooter=footer.GetPreferredSize(new Size(Math.Max(1,right),0)).Height>Math.Max(80,(root.ClientSize.Height-root.Padding.Vertical)*.35);
                if(scrollFooter!=footerInScroll)
                {
                    if(scrollFooter)Add(footer);
                    else {settings.Controls.Remove(footer);settings.RowStyles.RemoveAt(settings.RowStyles.Count-1);settings.RowCount=settings.Controls.Count;sidebar.Controls.Add(footer,0,1);}
                    footerInScroll=scrollFooter;
                }
                int contentWidth=Math.Max(minimumWidth,settingsViewport.ClientSize.Width-1);
                settings.Width=contentWidth;
                int bandWidth=contentWidth-settings.Padding.Horizontal;
                calibrationBand.ColumnStyles[1].Width=loupeWidth+magnifier.Margin.Horizontal;
                int columnWidth=bandWidth-loupeWidth-magnifier.Margin.Horizontal;scaleStatus.MaximumSize=new Size(Math.Max(1,columnWidth-scaleStatus.Margin.Horizontal),0);
                int calibrationHeight=calibrationValues.Controls.Cast<Control>().Sum(c=>c.GetPreferredSize(new Size(Math.Max(1,columnWidth-c.Margin.Horizontal),0)).Height+c.Margin.Vertical);
                calibrationValues.Height=calibrationHeight;calibrationBand.Height=Math.Max(calibrationHeight,magnifier.Height+magnifier.Margin.Vertical);
                int textWidth=Math.Max(1,bandWidth-12);status.MaximumSize=new Size(textWidth,0);navigationHint.MaximumSize=new Size(textWidth,0);
                int contentHeight=settings.Padding.Vertical+settings.Controls.Cast<Control>().Sum(c=>(c.AutoSize?c.GetPreferredSize(new Size(Math.Max(1,bandWidth-c.Margin.Horizontal),0)).Height:c.Height)+c.Margin.Vertical);
                settings.Bounds=new Rectangle(settingsViewport.AutoScrollPosition,new Size(contentWidth,contentHeight));settingsViewport.AutoScrollMinSize=new Size(contentWidth,contentHeight);root.PerformLayout();
            }
            finally{reflowing=false;}
        }
        private void FitScreen()
        {Rectangle work=Screen.FromHandle(Handle).WorkingArea;MinimumSize=new Size(Math.Min(MinimumSize.Width,work.Width-12),Math.Min(MinimumSize.Height,work.Height-12));Size=new Size((int)(work.Width*.96),(int)(work.Height*.96));Location=new Point(work.X+(work.Width-Width)/2,work.Y+(work.Height-Height)/2);}
        private void Export()
        {using(SaveFileDialog dialog=new SaveFileDialog {Filter="CSV 파일|*.csv",FileName="측정결과.csv"})if(dialog.ShowDialog(this)==DialogResult.OK)File.WriteAllText(dialog.FileName,Canvas.Document.Csv(),new UTF8Encoding(true));}
        private void Apply()
        {
            if(!Canvas.Document.HasScale)throw new InvalidOperationException("스케일을 먼저 설정하세요.");
            if(pendingCalibration!=null)throw new InvalidOperationException("새 기준에 스케일을 적용한 뒤 저장하세요.");
            if(Canvas.HasUnfinishedDrawing)throw new InvalidOperationException("그리는 도형을 먼저 완성하세요. 그리기 완료 또는 Esc로 취소한 뒤 적용할 수 있습니다.");
            apply.Enabled=false;UseWaitCursor=true;
            try{host.ApplyMeasurements(selection,Canvas.Document);DialogResult=DialogResult.OK;Close();}
            finally{UseWaitCursor=false;if(!IsDisposed)apply.Enabled=true;}
        }
        protected override void Dispose(bool disposing)
        {if(disposing){Canvas.Source=null;magnifier.Source=null;session.Dispose();tips.Dispose();foreach(Image image in buttonImages)image.Dispose();}base.Dispose(disposing);}
    }
    internal static class MeasurementIcons
    {
        internal static Bitmap Draw(string kind,int size)
        {
            Bitmap image=new Bitmap(size,size);using(Graphics g=Graphics.FromImage(image))using(Pen pen=new Pen(Color.FromArgb(40,75,110),2))
            {
                g.SmoothingMode=System.Drawing.Drawing2D.SmoothingMode.AntiAlias;g.ScaleTransform(size/32f,size/32f);
                if(kind=="ruler"){g.TranslateTransform(16,16);g.RotateTransform(-35);g.FillRectangle(Brushes.LightSteelBlue,-13,-6,26,12);g.DrawRectangle(pen,-13,-6,26,12);for(int x=-9;x<=10;x+=4)g.DrawLine(pen,x,-6,x,x%2==0?2:-1);}
                else if(kind=="circle"||kind=="circle3"||kind=="ellipse"){g.DrawEllipse(pen,4,kind=="ellipse"?9:4,24,kind=="ellipse"?14:24);if(kind=="circle")g.DrawLine(pen,16,16,27,16);else for(int i=0;i<3;i++){double t=i*2*Math.PI/3;g.FillEllipse(Brushes.SteelBlue,(float)(14+12*Math.Cos(t)),(float)(14+12*Math.Sin(t)),4,4);}}
                else if(kind=="rect"||kind=="rect3"){if(kind=="rect3"){g.TranslateTransform(16,16);g.RotateTransform(-20);g.TranslateTransform(-16,-16);}g.DrawRectangle(pen,5,7,23,18);}
                else if(kind=="angle"||kind=="hangle"||kind=="vangle"){g.DrawLines(pen,new[]{new Point(5,4),new Point(5,27),new Point(28,19)});g.DrawArc(pen,0,18,13,13,260,95);}
                else if(kind=="gap"||kind=="pointline"){g.DrawLine(pen,4,7,28,7);if(kind=="gap")g.DrawLine(pen,4,26,28,26);g.DrawLine(pen,16,7,16,26);g.FillEllipse(Brushes.SteelBlue,13,23,6,6);}
                else if(kind=="circle_distance"){g.DrawEllipse(pen,1,5,13,13);g.DrawEllipse(pen,18,14,13,13);g.DrawLine(pen,7,11,24,20);}
                else if(kind=="polygon"||kind=="lasso"||kind=="curve"){g.DrawPolygon(pen,new[]{new Point(4,8),new Point(18,3),new Point(27,14),new Point(21,28),new Point(8,24)});}
                else if(kind=="text"){using(Font f=new Font("Arial",25,FontStyle.Bold,GraphicsUnit.Pixel))g.DrawString("T",f,Brushes.SteelBlue,5,1);}
                else if(kind=="point"){g.DrawLine(pen,16,3,16,29);g.DrawLine(pen,3,16,29,16);}
                else if(kind=="polyline"||kind=="draw")g.DrawLines(pen,new[]{new Point(3,26),new Point(11,6),new Point(20,25),new Point(29,5)});
                else {g.DrawLine(pen,4,27,28,4);if(kind=="arrow")g.DrawLines(pen,new[]{new Point(17,5),new Point(28,4),new Point(27,15)});else{g.FillEllipse(Brushes.SteelBlue,1,24,6,6);g.FillEllipse(Brushes.SteelBlue,25,1,6,6);}}
            }return image;
        }
    }
}

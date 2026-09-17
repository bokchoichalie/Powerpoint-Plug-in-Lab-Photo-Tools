using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Windows.Forms;

namespace LabPhotoTools
{
    internal sealed partial class MeasurementCanvas : Control
    {
        internal MeasurementDocument Document;
        internal Image Source;
        internal double Rotation;
        internal MeasurementItem Style = new MeasurementItem();
        internal readonly HashSet<int> Selected = new HashSet<int>();
        internal string Tool = "select";
        internal bool Calibrating;
        internal event Action<MeasurementItem> CalibrationReady;
        internal event Action Changed;
        internal event Action<string> Status;
        internal event Action<MeasurePoint> HoverChanged;
        private readonly List<MeasurePoint> points = new List<MeasurePoint>();
        private readonly List<string> undo = new List<string>(), redo = new List<string>();
        private MeasurePoint hover = new MeasurePoint(), pan = new MeasurePoint(), dragStart;
        private Point lastMouse;
        private bool dragging, panning, freehand, dragLabel;
        private double zoom = 1;
        private int circleFirst;
        private Point? precisionMouse;
        internal MeasurePoint HoverPoint {get{return new MeasurePoint(hover.X,hover.Y);}}
        internal bool HasUnfinishedDrawing {get{return points.Count>0||circleFirst!=0||pointDragBefore!=null;}}
        private readonly Dictionary<int,List<MeasurePoint>> dragPoints = new Dictionary<int,List<MeasurePoint>>();
        private readonly Dictionary<int,MeasurePoint> dragLabels = new Dictionary<int,MeasurePoint>();
        private readonly Dictionary<int, RectangleF> labelBounds = new Dictionary<int, RectangleF>();
        private double Cos {get{return Math.Cos(Rotation*Math.PI/180);}}
        private double Sin {get{return Math.Sin(Rotation*Math.PI/180);}}
        internal double ViewScale { get { return Document == null ? 1 : Math.Max(.0001,Math.Min((ClientSize.Width-24)/(Math.Abs(Cos)*Document.Width+Math.Abs(Sin)*Document.Height),(ClientSize.Height-24)/(Math.Abs(Sin)*Document.Width+Math.Abs(Cos)*Document.Height)))*zoom; } }
        internal PointF ToScreen(MeasurePoint p)
        {double x=p.X-Document.Width/2,y=p.Y-Document.Height/2;return new PointF((float)(ClientSize.Width/2.0+pan.X+(x*Cos-y*Sin)*ViewScale),(float)(ClientSize.Height/2.0+pan.Y+(x*Sin+y*Cos)*ViewScale));}
        internal MeasurePoint ToImage(PointF p)
        {double x=(p.X-ClientSize.Width/2.0-pan.X)/ViewScale,y=(p.Y-ClientSize.Height/2.0-pan.Y)/ViewScale;return new MeasurePoint(Document.Width/2+x*Cos+y*Sin,Document.Height/2-x*Sin+y*Cos);}
        internal MeasurementCanvas()
        {
            DoubleBuffered=true;TabStop=true;BackColor=Color.FromArgb(28,34,43);Dock=DockStyle.Fill;
            SetStyle(ControlStyles.Selectable|ControlStyles.ResizeRedraw,true);
            AccessibleName="사진 측정 영역";
        }
        private void Tell(string s) { if(Status!=null)Status(s); }
        internal void ResetView() { zoom=1;pan=new MeasurePoint();Invalidate(); }
        internal void Zoom(double factor) { ZoomAt(new PointF(ClientSize.Width/2f,ClientSize.Height/2f),factor); }
        private void ZoomAt(PointF point,double factor)
        {
            if(Document==null)return;MeasurePoint before=ToImage(point);zoom=Math.Max(1,Math.Min(24,zoom*factor));
            PointF after=ToScreen(before);pan=pan+new MeasurePoint(point.X-after.X,point.Y-after.Y);Invalidate();
        }
        internal void SetTool(string tool,bool calibration)
        {
            CancelDrawing();Tool=tool;Calibrating=calibration;Cursor=tool=="select"?Cursors.Default:Cursors.Cross;Focus();Invalidate();
        }
        internal void CancelDrawing() { EndPointDrag(true);points.Clear();circleFirst=0;dragging=panning=freehand=false;Capture=false;Invalidate(); }
        internal void PushUndo()
        { undo.Add(Document.Serialize());if(undo.Count>100)undo.RemoveAt(0);redo.Clear(); }
        private void Notify() { Invalidate();if(Changed!=null)Changed(); }
        internal void Undo()
        {
            if(pointDragBefore!=null){EndPointDrag(true);return;}
            if(points.Count>0){points.RemoveAt(points.Count-1);Invalidate();return;}
            if(undo.Count==0)return;redo.Add(Document.Serialize());Document=MeasurementDocument.Deserialize(undo[undo.Count-1]);undo.RemoveAt(undo.Count-1);Selected.Clear();CancelDrawing();Notify();
        }
        internal void Redo()
        { if(pointDragBefore!=null){EndPointDrag(true);return;}if(redo.Count==0)return;undo.Add(Document.Serialize());Document=MeasurementDocument.Deserialize(redo[redo.Count-1]);redo.RemoveAt(redo.Count-1);Selected.Clear();CancelDrawing();Notify(); }
        internal void DeleteSelected()
        { if(Selected.Count==0)return;PushUndo();Document.Delete(Selected);Selected.Clear();Notify(); }
        internal void ClearAll() { if(Document.Items.Count==0)return;PushUndo();Document.Items.Clear();Selected.Clear();CancelDrawing();Notify(); }
        internal void SelectAll() { Selected.Clear();foreach(MeasurementItem i in Document.Items)Selected.Add(i.Id);Notify(); }
        internal void ApplyStyle()
        {
            if(Selected.Count==0)return;PushUndo();
            foreach(MeasurementItem i in Document.Items.Where(i=>Selected.Contains(i.Id)))
            { i.ColorArgb=Style.ColorArgb;i.TextColorArgb=Style.TextColorArgb;i.LineWidth=Style.LineWidth;i.FontSize=Style.FontSize;i.Dashed=Style.Dashed;i.Filled=Style.Filled; }
            Notify();
        }
        private MeasurementItem NewItem()
        {
            return new MeasurementItem {Kind=Tool,Points=points.Select(p=>new MeasurePoint(p.X,p.Y)).ToList(),ColorArgb=Style.ColorArgb,TextColorArgb=Style.TextColorArgb,LineWidth=Style.LineWidth,FontSize=Style.FontSize,Dashed=Style.Dashed,Guide=Style.Guide,Filled=Style.Filled,Obtuse=Style.Obtuse,Note=Style.Note};
        }
        internal void Finish()
        {
            if(Tool=="editpoints"){EndPointDrag(false);SetTool("select",false);Notify();return;}
            if(points.Count==0)return;
            try
            {
                MeasurementItem item=NewItem();
                if(Calibrating)
                {
                    double length=MeasurementGeometry.ReferenceLength(item);
                    if(length<.01)throw new InvalidOperationException("기준 길이가 너무 짧습니다. 다시 지정하세요.");
                    if(CalibrationReady!=null)CalibrationReady(item);
                    points.Clear();Tool="select";Calibrating=false;Notify();return;
                }
                MeasurementGeometryResult result=MeasurementGeometry.Build(item,Document);
                if(Document.Items.Count>=1000)throw new InvalidOperationException("한 사진에는 최대 1,000개 측정 도형을 저장할 수 있습니다.");
                PushUndo();item.Id=Document.NextId++;Document.Items.Add(item);points.Clear();Selected.Clear();Selected.Add(item.Id);
                Tell(MeasurementGeometry.Names[item.Kind]+": "+result.Text.Replace("\n"," / "));Notify();
            }
            catch(Exception ex){Tell(ex.Message);}
        }
        internal void ClickImage(MeasurePoint p)
        {
            if(Document==null||p.X<0||p.Y<0||p.X>Document.Width||p.Y>Document.Height)return;
            if(Tool=="select"||Tool=="erase"||Tool=="editpoints")return;
            if(Tool=="circle_distance")
            {
                MeasurementItem found=Hit(p,true);
                if(found==null){Tell("기존 원의 둘레 또는 중심을 클릭하세요.");return;}
                if(circleFirst==0){circleFirst=found.Id;Selected.Clear();Selected.Add(found.Id);Tell("두 번째 원을 클릭하세요.");Notify();return;}
                if(circleFirst==found.Id){Tell("서로 다른 원을 선택하세요.");return;}
                MeasurementItem item=NewItem();item.CircleA=circleFirst;item.CircleB=found.Id;
                MeasurementGeometry.Build(item,Document);PushUndo();item.Id=Document.NextId++;Document.Items.Add(item);circleFirst=0;Selected.Clear();Selected.Add(item.Id);Notify();return;
            }
            if(!Calibrating&&!Document.HasScale&&Tool!="select"&&Tool!="erase") { Tell("먼저 사진의 스케일바로 스케일을 맞추세요.");return; }
            if(points.Count>=(Tool=="ncurve"?512:5000)){Tell("지정 가능한 점 수에 도달했습니다. Enter로 완성하세요.");return;}
            points.Add(p);
            int n=Calibrating?(Tool=="line"?2:3):MeasurementGeometry.PointCount(Tool);
            if(n>0&&points.Count>=n)Finish();else Invalidate();
        }
        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);Focus();if(Document==null)return;lastMouse=e.Location;
            if(e.Button==MouseButtons.Middle){panning=true;Capture=true;return;}
            if(e.Button==MouseButtons.Right){if(points.Count>0)Finish();else {Tool="select";Notify();}return;}
            if(e.Button!=MouseButtons.Left)return;
            MeasurePoint p=precisionMouse.HasValue&&e.Location==precisionMouse.Value?HoverPoint:ToImage(e.Location);hover=p;
            if(Tool=="editpoints"){StartPointDrag(p);return;}
            if(Tool=="select"||Tool=="erase")
            {
                MeasurementItem item=Hit(p,false);
                if(item==null){Selected.Clear();Notify();return;}
                if(Tool=="erase"){PushUndo();Document.Delete(new[]{item.Id});Selected.Clear();Notify();return;}
                if((ModifierKeys&Keys.Control)!=0){if(!Selected.Add(item.Id))Selected.Remove(item.Id);Notify();return;}
                if(!Selected.Contains(item.Id)){Selected.Clear();Selected.Add(item.Id);}
                dragLabel=labelBounds.ContainsKey(item.Id)&&labelBounds[item.Id].Contains(e.Location);
                PushUndo();dragStart=p;dragging=true;Capture=true;dragPoints.Clear();dragLabels.Clear();
                foreach(MeasurementItem s in Document.Items.Where(s=>Selected.Contains(s.Id)))
                {dragPoints[s.Id]=s.Points.Select(q=>new MeasurePoint(q.X,q.Y)).ToList();dragLabels[s.Id]=new MeasurePoint(s.LabelOffset.X,s.LabelOffset.Y);}
                Notify();return;
            }
            if(Tool=="lasso"||Tool=="draw")
            {if(!Document.HasScale){Tell("먼저 스케일을 맞추세요.");return;}points.Clear();freehand=true;Capture=true;}
            ClickImage(p);
        }
        protected override void OnMouseEnter(EventArgs e)
        {base.OnMouseEnter(e);Focus();}
        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);if(Document==null)return;
            if(panning){pan=pan+new MeasurePoint(e.X-lastMouse.X,e.Y-lastMouse.Y);lastMouse=e.Location;Invalidate();return;}
            // Cursor.Position produces an integer screen coordinate. Retain the
            // exact sub-screen-pixel image coordinate until the mouse really moves.
            if(precisionMouse.HasValue&&e.Location==precisionMouse.Value)return;
            precisionMouse=null;hover=ToImage(e.Location);
            if(pointDragBefore!=null){MoveEditedPoint(hover);if(HoverChanged!=null)HoverChanged(hover);return;}
            if(dragging)
            {
                MeasurePoint d=hover-dragStart;
                foreach(MeasurementItem item in Document.Items.Where(i=>Selected.Contains(i.Id)))
                {
                    if(dragLabel||item.Kind=="circle_distance")item.LabelOffset=dragLabels[item.Id]+d;
                    else item.Points=dragPoints[item.Id].Select(p=>p+d).ToList();
                }
                Invalidate();return;
            }
            if(freehand&&points.Count<5000&&MeasurementGeometry.Distance(points[points.Count-1],hover)>2/ViewScale)
                points.Add(new MeasurePoint(Math.Max(0,Math.Min(Document.Width,hover.X)),Math.Max(0,Math.Min(Document.Height,hover.Y))));
            if(HoverChanged!=null)HoverChanged(hover);Invalidate();
        }
        protected override void OnMouseUp(MouseEventArgs e)
        {base.OnMouseUp(e);if(pointDragBefore!=null)EndPointDrag(false);if(freehand){freehand=false;Finish();}if(dragging)Notify();panning=dragging=false;Capture=false;}
        protected override void OnMouseWheel(MouseEventArgs e) {base.OnMouseWheel(e);ZoomAt(e.Location,e.Delta>0?1.25:.8);}
        protected override bool IsInputKey(Keys keyData)
        {Keys k=keyData&Keys.KeyCode;if(k==Keys.Left||k==Keys.Right||k==Keys.Up||k==Keys.Down||k==Keys.Enter||k==Keys.Escape)return true;return base.IsInputKey(keyData);}
        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);if(Document==null)return;e.Handled=true;
            if(e.Control&&e.KeyCode==Keys.Z){Undo();return;}if(e.Control&&e.KeyCode==Keys.Y){Redo();return;}
            if(e.Control&&e.KeyCode==Keys.A){SelectAll();return;}
            if(e.KeyCode==Keys.Delete){if(e.Shift)ClearAll();else DeleteSelected();return;}
            if(e.KeyCode==Keys.Escape){CancelDrawing();Tool="select";Notify();return;}
            if(e.KeyCode==Keys.Enter){if(Tool=="editpoints"||points.Count>0&&(Tool=="ncurve"||Tool=="polygon"||Tool=="curve"||Tool=="polyline"||Tool=="gap"||Tool=="pointline"))Finish();else ClickImage(new MeasurePoint(hover.X,hover.Y));return;}
            if(e.Control||e.Alt){e.Handled=false;return;}
            double step=e.Shift?10:1,dx=0,dy=0;
            if(e.KeyCode==Keys.A||e.KeyCode==Keys.Left)dx=-step;else if(e.KeyCode==Keys.D||e.KeyCode==Keys.Right)dx=step;
            else if(e.KeyCode==Keys.W||e.KeyCode==Keys.Up)dy=-step;else if(e.KeyCode==Keys.S||e.KeyCode==Keys.Down)dy=step;
            else {e.Handled=false;return;}
            NudgeCursor(dx,dy);e.SuppressKeyPress=true;
        }
        internal void NudgeCursor(double dx,double dy)
        {
            if(Document==null||Source==null)return;
            // One pixel of the loaded image, independent of zoom and a saved
            // document's coordinate extent. Directions follow the visible photo.
            double px=Document.Width/Source.Width,py=Document.Height/Source.Height;
            hover=new MeasurePoint(Math.Max(0,Math.Min(Document.Width-px,hover.X+(dx*Cos+dy*Sin)*px)),Math.Max(0,Math.Min(Document.Height-py,hover.Y+(-dx*Sin+dy*Cos)*py)));
            PointF screen=ToScreen(hover);
            // Keep a nudged point in view even when zoomed into an image edge.
            float sx=Math.Max(8,Math.Min(ClientSize.Width-9,screen.X)),sy=Math.Max(8,Math.Min(ClientSize.Height-9,screen.Y));
            pan=pan+new MeasurePoint(sx-screen.X,sy-screen.Y);screen=ToScreen(hover);
            precisionMouse=System.Drawing.Point.Round(screen);
            if(IsHandleCreated&&Focused)System.Windows.Forms.Cursor.Position=PointToScreen(precisionMouse.Value);
            if(HoverChanged!=null)HoverChanged(hover);Invalidate();
        }
        private MeasurementItem Hit(MeasurePoint p,bool circlesOnly)
        {
            PointF screen=ToScreen(p);MeasurementItem best=null;double closest=12/ViewScale;
            foreach(MeasurementItem item in Document.Items.AsEnumerable().Reverse())
            {
                if(circlesOnly&&!MeasurementGeometry.IsCircle(item))continue;
                if(!circlesOnly&&labelBounds.ContainsKey(item.Id)&&labelBounds[item.Id].Contains(screen))return item;
                MeasurementGeometryResult r=MeasurementGeometry.Build(item,Document);
                double d=double.MaxValue;
                foreach(MeasurementPath path in r.Paths)
                {for(int i=1;i<path.Points.Count;i++)d=Math.Min(d,MeasurementGeometry.SegmentDistance(p,path.Points[i-1],path.Points[i]));if(path.Closed)d=Math.Min(d,MeasurementGeometry.SegmentDistance(p,path.Points[path.Points.Count-1],path.Points[0]));}
                if(circlesOnly&&r.Center!=null)d=Math.Min(d,MeasurementGeometry.Distance(p,r.Center));
                if(d<=closest){closest=d;best=item;}
            }
            return best;
        }
        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);if(Document==null||Source==null)return;
            Graphics g=e.Graphics;g.SmoothingMode=SmoothingMode.AntiAlias;g.InterpolationMode=zoom>3?InterpolationMode.NearestNeighbor:InterpolationMode.HighQualityBicubic;
            // Draw the exported local image through the same transform as the
            // measurement points. PowerPoint's flips are already in the pixels.
            g.DrawImage(Source,new[]{ToScreen(new MeasurePoint(0,0)),ToScreen(new MeasurePoint(Document.Width,0)),ToScreen(new MeasurePoint(0,Document.Height))});
            labelBounds.Clear();
            foreach(MeasurementItem item in Document.Items)DrawItem(g,item,Selected.Contains(item.Id),false);
            if(Tool=="editpoints")DrawEditPoints(g);
            if(points.Count>0)
            {
                MeasurementItem draft=NewItem();draft.Points.Add(hover);
                try {DrawItem(g,draft,false,true);} catch { }
                using(Pen pen=new Pen(Color.Cyan,1.5f)){if(points.Count>1)g.DrawLines(pen,points.Select(ToScreen).ToArray());foreach(MeasurePoint p in points){PointF q=ToScreen(p);g.DrawEllipse(pen,q.X-3,q.Y-3,6,6);}}
            }
            if(Focused&&Tool!="select"&&Tool!="erase")
            {PointF h=ToScreen(hover);using(Pen pen=new Pen(Color.Cyan,1)){g.DrawLine(pen,h.X-8,h.Y,h.X+8,h.Y);g.DrawLine(pen,h.X,h.Y-8,h.X,h.Y+8);}}
        }
        private void DrawItem(Graphics g,MeasurementItem item,bool selected,bool preview)
        {
            MeasurementGeometryResult r=MeasurementGeometry.Build(item,Document);double units=Document.Width/800;
            using(Pen pen=new Pen(preview?Color.Cyan:Color.FromArgb(item.ColorArgb),Math.Max(1,(float)(item.LineWidth*units*ViewScale))))
            {
                pen.DashStyle=item.Dashed||preview?DashStyle.Dash:DashStyle.Solid;
                foreach(MeasurementPath path in r.Paths)
                {
                    PointF[] pts=path.Points.Select(ToScreen).ToArray();if(pts.Length<2)continue;
                    if(selected)using(Pen glow=new Pen(Color.FromArgb(150,80,160,255),pen.Width+5)){if(path.Closed)g.DrawPolygon(glow,pts);else g.DrawLines(glow,pts);}
                    if(path.Closed){if(item.Filled)using(Brush fill=new SolidBrush(Color.FromArgb(60,Color.FromArgb(item.ColorArgb))))g.FillPolygon(fill,pts);g.DrawPolygon(pen,pts);}else g.DrawLines(pen,pts);
                }
            }
            if(string.IsNullOrEmpty(r.Text)||preview)return;
            string text=(item.Guide||item.Kind=="text"?"":"#"+item.Id+" ")+r.Text;
            PointF loc=ToScreen(r.Label);float fontPixels=Math.Max(2,(float)(item.FontSize*units*ViewScale));
            using(Font font=new Font("Arial",fontPixels,FontStyle.Regular,GraphicsUnit.Pixel))
            using(Brush brush=new SolidBrush(Color.FromArgb(item.TextColorArgb)))
            {
                SizeF size=g.MeasureString(text,font);RectangleF box=new RectangleF(loc.X+5,loc.Y+5,size.Width+6,size.Height+4);
                using(Brush back=new SolidBrush(Color.FromArgb(175,20,24,30)))g.FillRectangle(back,box);
                g.DrawString(text,font,brush,box.X+3,box.Y+2);labelBounds[item.Id]=box;
            }
        }
    }
    internal sealed class MeasurementMagnifier : Control
    {
        internal Image Source;internal MeasurePoint Point;internal MeasurementDocument Document;internal double Rotation;
        internal MeasurementMagnifier(){DoubleBuffered=true;BackColor=Color.FromArgb(28,34,43);}
        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);if(Source==null||Point==null||Document==null)return;
            double sx=Source.Width/Document.Width,sy=Source.Height/Document.Height;
            Graphics g=e.Graphics;GraphicsState saved=g.Save();g.TranslateTransform(ClientSize.Width/2f,ClientSize.Height/2f);g.RotateTransform((float)Rotation);g.ScaleTransform(6,6);g.TranslateTransform((float)(-Point.X*sx),(float)(-Point.Y*sy));
            g.InterpolationMode=InterpolationMode.NearestNeighbor;g.PixelOffsetMode=PixelOffsetMode.Half;g.DrawImage(Source,new Rectangle(0,0,Source.Width,Source.Height));g.Restore(saved);
            e.Graphics.DrawLine(Pens.Cyan,Width/2,0,Width/2,Height);e.Graphics.DrawLine(Pens.Cyan,0,Height/2,Width,Height/2);
        }
    }
}

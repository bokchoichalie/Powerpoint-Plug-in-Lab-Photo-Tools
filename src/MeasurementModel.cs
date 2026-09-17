using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Web.Script.Serialization;

namespace LabPhotoTools
{
    // Manual calibration and geometry follow Prusacope v19's MeasurePanel,
    // MeasureToolsPanel, _circle_from_3points and _on_measure_shape_selected.
    // Coordinates belong to the calibration image, never to the zoomed view.
    public sealed class MeasurePoint
    {
        public double X, Y;
        public MeasurePoint() { }
        public MeasurePoint(double x, double y) { X = x; Y = y; }
        public static MeasurePoint operator +(MeasurePoint a, MeasurePoint b) { return new MeasurePoint(a.X + b.X, a.Y + b.Y); }
        public static MeasurePoint operator -(MeasurePoint a, MeasurePoint b) { return new MeasurePoint(a.X - b.X, a.Y - b.Y); }
        public static MeasurePoint operator *(MeasurePoint a, double b) { return new MeasurePoint(a.X * b, a.Y * b); }
    }
    public sealed class MeasurementItem
    {
        public int Id;
        public string Kind = "line", Note = "";
        public List<MeasurePoint> Points = new List<MeasurePoint>();
        public List<MeasurePoint> Tangents = new List<MeasurePoint>();
        public int CircleA, CircleB;
        public int ColorArgb = Color.Yellow.ToArgb(), TextColorArgb = Color.White.ToArgb();
        public float LineWidth = 2, FontSize = 13;
        public bool Dashed, Guide, Filled, Obtuse;
        public MeasurePoint LabelOffset = new MeasurePoint();
    }
    public sealed class MeasurementDocument
    {
        public int Version = 1, NextId = 1;
        public double Width, Height, MicronsPerPixel;
        public string Unit = "µm", PhotoSignature = "";
        public MeasurementItem Calibration;
        public double CalibrationLengthMicrons;
        public List<MeasurementItem> Items = new List<MeasurementItem>();
        public bool HasScale { get { return MeasurementGeometry.Finite(MicronsPerPixel) && MicronsPerPixel > 0; } }
        public void Calibrate(MeasurementItem reference, double realLength, string unit)
        {
            double pixels = MeasurementGeometry.ReferenceLength(reference);
            double microns = realLength * MeasurementGeometry.UnitFactor(unit);
            if (!MeasurementGeometry.Finite(microns) || microns <= 0 || pixels < .01)
                throw new InvalidOperationException("기준 길이가 0입니다. 서로 다른 점을 지정하고 실제 길이를 입력하세요.");
            Calibration = reference; CalibrationLengthMicrons = microns;
            MicronsPerPixel = microns / pixels; Unit = unit;
        }
        public string Serialize() { return new JavaScriptSerializer { MaxJsonLength = 4000000 }.Serialize(this); }
        public static MeasurementDocument Deserialize(string json)
        {
            MeasurementDocument d = new JavaScriptSerializer { MaxJsonLength = 4000000 }.Deserialize<MeasurementDocument>(json);
            if (d == null || d.Version != 1 || !MeasurementGeometry.Finite(d.Width) || !MeasurementGeometry.Finite(d.Height) || d.Width <= 0 || d.Height <= 0 || d.Items == null || d.Items.Count > 1000)
                throw new InvalidOperationException("저장된 측정 정보를 읽을 수 없습니다.");
            MeasurementGeometry.UnitFactor(d.Unit);
            if (d.MicronsPerPixel != 0 && !d.HasScale) throw new InvalidOperationException("저장된 스케일이 올바르지 않습니다.");
            HashSet<int> ids=new HashSet<int>();
            foreach (MeasurementItem item in d.Items)
            {
                if (item == null || item.Kind==null || !MeasurementGeometry.Names.ContainsKey(item.Kind) || item.Points == null || item.Points.Count > 5000 || item.LabelOffset == null || !MeasurementGeometry.Finite(item.LabelOffset.X) || !MeasurementGeometry.Finite(item.LabelOffset.Y) || item.Id<=0 || !ids.Add(item.Id) || !MeasurementGeometry.Finite(item.LineWidth) || item.LineWidth<1 || item.LineWidth>12 || !MeasurementGeometry.Finite(item.FontSize) || item.FontSize<6 || item.FontSize>72 || item.Points.Any(p => p == null || !MeasurementGeometry.Finite(p.X) || !MeasurementGeometry.Finite(p.Y)))
                    throw new InvalidOperationException("저장된 측정 도형이 올바르지 않습니다.");
                if(item.Tangents==null)item.Tangents=new List<MeasurePoint>();
                if(item.Tangents.Count!=0&&(item.Kind!="ncurve"||item.Tangents.Count!=item.Points.Count||item.Tangents.Any(p=>p==null||!MeasurementGeometry.Finite(p.X)||!MeasurementGeometry.Finite(p.Y))))
                    throw new InvalidOperationException("저장된 곡률 정보가 올바르지 않습니다.");
                MeasurementGeometry.Build(item, d);
            }
            d.NextId=Math.Max(d.NextId,d.Items.Count==0?1:d.Items.Max(i=>i.Id)+1);
            return d;
        }
        public void Delete(IEnumerable<int> ids)
        {
            HashSet<int> removed = new HashSet<int>(ids);
            Items.RemoveAll(i => removed.Contains(i.Id) || (i.Kind == "circle_distance" && (removed.Contains(i.CircleA) || removed.Contains(i.CircleB))));
        }
        public string Csv()
        {
            StringBuilder b = new StringBuilder("ID,Tool,Result,Scale (um/pixel)\r\n");
            foreach (MeasurementItem i in Items.Where(i => !i.Guide && i.Kind != "text" && i.Kind != "draw" && i.Kind != "arrow"))
                b.Append(i.Id).Append(",\"").Append(MeasurementGeometry.Names[i.Kind]).Append("\",\"").Append(MeasurementGeometry.Build(i, this).Text.Replace("\"", "\"\"").Replace("\n", "; ")).Append("\",").Append(MicronsPerPixel.ToString("R", CultureInfo.InvariantCulture)).Append("\r\n");
            return b.ToString();
        }
    }
    public sealed class MeasurementPath
    {
        public List<MeasurePoint> Points = new List<MeasurePoint>();
        public bool Closed;
    }
    public sealed class MeasurementGeometryResult
    {
        public List<MeasurementPath> Paths = new List<MeasurementPath>();
        public string Text = "";
        public MeasurePoint Label = new MeasurePoint(), Center;
        public double Radius, Length, Area, Width, Height, Angle;
    }
    public static class MeasurementGeometry
    {
        public static readonly Dictionary<string, string> Names = new Dictionary<string, string> {
            {"line", "선"}, {"angle", "3점 각도"}, {"hangle", "수평 각도"}, {"vangle", "수직 각도"},
            {"rect", "사각형"}, {"rect3", "3점 사각형"}, {"circle", "중심·반지름 원"}, {"circle3", "3점원"},
            {"circle_distance", "원 사이 거리"}, {"ellipse", "3점 타원"}, {"polygon", "다각형 면적"},
            {"polyline", "꺾은선 길이"}, {"lasso", "자유곡선 면적"}, {"curve", "곡선 면적"}, {"gap", "평행선 간격"},
            {"pointline", "점·직선 거리"}, {"point", "점"}, {"text", "텍스트"}, {"arrow", "화살표"}, {"draw", "자유선"}
            ,{"ncurve", "n점원"}
        };
        public static bool Finite(double v) { return !double.IsNaN(v) && !double.IsInfinity(v); }
        public static double UnitFactor(string unit)
        {
            switch (unit) { case "nm": return .001; case "µm": return 1; case "mm": return 1000; case "cm": return 10000; default: throw new ArgumentException("단위를 선택하세요."); }
        }
        public static int PointCount(string kind)
        {
            if (kind == "point" || kind == "text") return 1;
            if (kind == "ncurve" || kind == "polygon" || kind == "polyline" || kind == "lasso" || kind == "curve" || kind == "draw" || kind == "gap" || kind == "pointline") return 0;
            if (kind == "angle" || kind == "rect3" || kind == "circle3" || kind == "ellipse") return 3;
            return 2;
        }
        public static double Distance(MeasurePoint a, MeasurePoint b) { return Math.Sqrt((a.X-b.X)*(a.X-b.X)+(a.Y-b.Y)*(a.Y-b.Y)); }
        private static double Dot(MeasurePoint a, MeasurePoint b) { return a.X*b.X+a.Y*b.Y; }
        private static void Positive(double v) { if (!Finite(v) || v < .000001) throw new InvalidOperationException("점이 겹치거나 도형의 크기가 0입니다. 점을 다시 지정하세요."); }
        public static MeasurePoint Project(MeasurePoint p, MeasurePoint a, MeasurePoint b)
        { MeasurePoint v = b-a; double n = Dot(v,v); Positive(n); return a + v * (Dot(p-a,v)/n); }
        public static double PolygonArea(IList<MeasurePoint> p)
        { double v=0; for(int i=0;i<p.Count;i++){MeasurePoint a=p[i],b=p[(i+1)%p.Count]; v+=a.X*b.Y-b.X*a.Y;} return Math.Abs(v)*.5; }
        public static double Perimeter(IList<MeasurePoint> p, bool closed)
        { double v=0;for(int i=1;i<p.Count;i++)v+=Distance(p[i-1],p[i]);if(closed&&p.Count>2)v+=Distance(p[p.Count-1],p[0]);return v; }
        public static MeasurePoint CircleCenter(IList<MeasurePoint> p)
        {
            // Translate first: avoids cancellation for small circles far from (0,0).
            MeasurePoint a=p[1]-p[0],b=p[2]-p[0]; double d=2*(a.X*b.Y-a.Y*b.X);
            if(Math.Abs(d)<1e-8*Math.Max(1,Distance(p[0],p[1])*Distance(p[0],p[2])))
                throw new InvalidOperationException("세 점이 일직선에 가깝습니다. 원 둘레에서 다른 점을 선택하세요.");
            return p[0]+new MeasurePoint((Dot(a,a)*b.Y-Dot(b,b)*a.Y)/d,(a.X*Dot(b,b)-b.X*Dot(a,a))/d);
        }
        public static double ReferenceLength(MeasurementItem item)
        {
            if(item == null || item.Points == null || item.Points.Count < 2) throw new InvalidOperationException("기준선을 먼저 지정하세요.");
            if(item.Kind=="circle3") { if(item.Points.Count<3)throw new InvalidOperationException("원 둘레의 세 점을 지정하세요.");return 2*Distance(CircleCenter(item.Points),item.Points[0]); }
            if(item.Kind=="gap") { if(item.Points.Count<3)throw new InvalidOperationException("평행선 위치를 지정하세요.");return Distance(item.Points[2],Project(item.Points[2],item.Points[0],item.Points[1])); }
            return Distance(item.Points[0],item.Points[1]);
        }
        private static string Number(double v) { return v.ToString(v!=0&&Math.Abs(v)<.001?"0.###E+0":"0.###", CultureInfo.InvariantCulture); }
        private static string Length(double px, MeasurementDocument d) { return Number(px*d.MicronsPerPixel/UnitFactor(d.Unit))+" "+d.Unit; }
        private static string Area(double px, MeasurementDocument d) { double s=d.MicronsPerPixel/UnitFactor(d.Unit);return Number(px*s*s)+" "+d.Unit+"²"; }
        private static void Path(MeasurementGeometryResult r, bool closed, params MeasurePoint[] pts) { r.Paths.Add(new MeasurementPath { Closed=closed, Points=pts.ToList() }); }
        public static MeasurementGeometryResult Build(MeasurementItem item, MeasurementDocument doc)
        {
            MeasurementGeometryResult r=new MeasurementGeometryResult(); List<MeasurePoint> p=item.Points;
            string k=item.Kind; int needed=PointCount(k);
            if(k!="circle_distance" && p.Count<Math.Max(needed, k=="polygon"||k=="lasso"||k=="curve"||k=="gap"||k=="pointline"?3:2) && k!="point" && k!="text")
                throw new InvalidOperationException("도형을 완성할 점을 더 지정하세요.");
            if(k=="circle_distance")
            {
                MeasurementItem first=doc.Items.Find(x=>x.Id==item.CircleA), second=doc.Items.Find(x=>x.Id==item.CircleB);
                if(first==null||second==null||first.Id==second.Id||!IsCircle(first)||!IsCircle(second))throw new InvalidOperationException("서로 다른 원 두 개를 선택하세요.");
                MeasurementGeometryResult a=Build(first,doc), b=Build(second,doc);
                r.Length=Distance(a.Center,b.Center); r.Width=Math.Max(0,r.Length-a.Radius-b.Radius);
                Path(r,false,a.Center,b.Center); r.Label=(a.Center+b.Center)*.5;
                r.Text="Min="+Length(r.Width,doc)+"\nCenter="+Length(r.Length,doc);
            }
            else if(k=="circle"||k=="circle3"||k=="ellipse")
            {
                double rx,ry,theta=0;
                if(k=="ellipse")
                { r.Center=(p[0]+p[1])*.5;rx=Distance(p[0],p[1])*.5;ry=Distance(p[2],Project(p[2],p[0],p[1]));theta=Math.Atan2(p[1].Y-p[0].Y,p[1].X-p[0].X); }
                else { r.Center=k=="circle"?p[0]:CircleCenter(p);rx=ry=Distance(r.Center,k=="circle"?p[1]:p[0]); }
                Positive(rx);Positive(ry);r.Radius=rx;r.Width=2*rx;r.Height=2*ry;r.Area=Math.PI*rx*ry;
                List<MeasurePoint> ring=new List<MeasurePoint>();
                for(int i=0;i<120;i++){double t=i*Math.PI/60,x=rx*Math.Cos(t),y=ry*Math.Sin(t);ring.Add(r.Center+new MeasurePoint(x*Math.Cos(theta)-y*Math.Sin(theta),x*Math.Sin(theta)+y*Math.Cos(theta)));}
                Path(r,true,ring.ToArray());r.Label=r.Center;
                r.Text=(k=="ellipse"?"W="+Length(r.Width,doc)+"  H="+Length(r.Height,doc):"D="+Length(2*rx,doc)+"  R="+Length(rx,doc))+"\nA="+Area(r.Area,doc);
                if(k=="circle")Path(r,false,r.Center,p[1]);
            }
            else if(k=="rect"||k=="rect3")
            {
                MeasurePoint a=p[0],b=p[1],c,e;
                if(k=="rect") { a=new MeasurePoint(Math.Min(p[0].X,p[1].X),Math.Min(p[0].Y,p[1].Y));c=new MeasurePoint(Math.Max(p[0].X,p[1].X),Math.Max(p[0].Y,p[1].Y));b=new MeasurePoint(c.X,a.Y);e=new MeasurePoint(a.X,c.Y); }
                else { MeasurePoint delta=p[2]-Project(p[2],a,b);c=b+delta;e=a+delta; }
                r.Width=Distance(a,b);r.Height=Distance(b,c);Positive(r.Width);Positive(r.Height);r.Area=r.Width*r.Height;
                Path(r,true,a,b,c,e);r.Label=(a+c)*.5;r.Text="W="+Length(r.Width,doc)+"  H="+Length(r.Height,doc)+"\nA="+Area(r.Area,doc);
            }
            else if(k=="angle"||k=="hangle"||k=="vangle")
            {
                MeasurePoint vertex=k=="angle"?p[1]:p[0],u,v;
                if(k=="angle"){u=p[0]-vertex;v=p[2]-vertex;Path(r,false,p[0],vertex,p[2]);}
                else
                {
                    v=p[1]-vertex; u=k=="hangle"?new MeasurePoint(v.X<0?-1:1,0):new MeasurePoint(0,v.Y<0?-1:1);
                    if(item.Obtuse)u=u*-1;u=u*Distance(vertex,p[1]);Path(r,false,vertex+u,vertex,p[1]);
                }
                double nu=Math.Sqrt(Dot(u,u)),nv=Math.Sqrt(Dot(v,v));Positive(nu);Positive(nv);
                r.Angle=Math.Acos(Math.Max(-1,Math.Min(1,Dot(u,v)/(nu*nv))))*180/Math.PI;r.Label=vertex;r.Text=Number(r.Angle)+"°";
            }
            else if(k=="gap"||k=="pointline")
            {
                Path(r,false,p[0],p[1]);List<string> labels=new List<string>();
                for(int i=2;i<p.Count;i++)
                {
                    MeasurePoint foot=Project(p[i],p[0],p[1]);double distance=Distance(p[i],foot);
                    if(k=="gap") { MeasurePoint shift=p[i]-foot;Path(r,false,p[0]+shift,p[1]+shift); }
                    Path(r,false,foot,p[i]);labels.Add(Length(distance,doc));r.Length=distance;
                }
                r.Label=p[p.Count-1];r.Text=string.Join("\n",labels);
            }
            else if(k=="point"||k=="text")
            {
                if(p.Count<1)throw new InvalidOperationException("위치를 지정하세요.");
                r.Label=p[0];
                if(k=="point"){double s=Math.Max(doc.Width,doc.Height)/250;Path(r,false,p[0]+new MeasurePoint(-s,0),p[0]+new MeasurePoint(s,0));Path(r,false,p[0]+new MeasurePoint(0,-s),p[0]+new MeasurePoint(0,s));r.Text="X="+Length(p[0].X,doc)+" Y="+Length(p[0].Y,doc);}
                else r.Text=item.Note??"";
            }
            else
            {
                if(k=="ncurve")p=MeasurementSpline.Sample(item);
                if(k=="curve")
                {
                    List<MeasurePoint> samples=new List<MeasurePoint>();
                    for(int n=0;n+2<p.Count;n+=2)for(int j=0;j<24;j++){double t=j/24.0;samples.Add(p[n]*((1-t)*(1-t))+p[n+1]*(2*(1-t)*t)+p[n+2]*(t*t));}
                    samples.Add(p[p.Count-1]);p=samples;
                }
                bool closed=k=="ncurve"||k=="polygon"||k=="lasso"||k=="curve";r.Length=Perimeter(p,closed);Positive(r.Length);Path(r,closed,p.ToArray());
                r.Area=closed?PolygonArea(p):0;r.Label=p[p.Count/2];r.Text=closed?"A="+Area(r.Area,doc)+"\nP="+Length(r.Length,doc):Length(r.Length,doc);
                if(k=="ncurve"){Positive(r.Area);r.Label=new MeasurePoint(item.Points.Average(q=>q.X),item.Points.Average(q=>q.Y));}
                if(k=="arrow") { MeasurePoint v=p[0]-p[1];double n=Distance(p[0],p[1]);v=v*(Math.Min(n*.25,doc.Width*.018)/n);MeasurePoint normal=new MeasurePoint(-v.Y,v.X)*.5;Path(r,false,p[1]+v+normal,p[1],p[1]+v-normal); }
                if(k=="arrow"||k=="draw")r.Text="";
            }
            if(item.Guide)r.Text="";
            r.Label=r.Label+(item.LabelOffset??new MeasurePoint());
            return r;
        }
        public static bool IsCircle(MeasurementItem i) { return i.Kind=="circle"||i.Kind=="circle3"; }
        public static double SegmentDistance(MeasurePoint p,MeasurePoint a,MeasurePoint b)
        { double n=Dot(b-a,b-a);if(n<1e-12)return Distance(p,a);double t=Math.Max(0,Math.Min(1,Dot(p-a,b-a)/n));return Distance(p,a+(b-a)*t); }
    }
}

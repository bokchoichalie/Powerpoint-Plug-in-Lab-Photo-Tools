using System;
using System.Collections.Generic;
using System.Linq;

namespace LabPhotoTools
{
    // Interpolating closed cubic Beziers. Each anchor has a symmetric tangent
    // vector: incoming P-T and outgoing P+T, so adjacent segments meet smoothly.
    internal static class MeasurementSpline
    {
        internal static MeasurePoint AutoTangent(IList<MeasurePoint> points,int index)
        {
            int n=points.Count;MeasurePoint previous=points[(index+n-1)%n],next=points[(index+1)%n],p=points[index];
            MeasurePoint direction=next-previous;double length=MeasurementGeometry.Distance(next,previous);
            double reach=Math.Min(MeasurementGeometry.Distance(p,previous),MeasurementGeometry.Distance(p,next))/3;
            return length<1e-10?new MeasurePoint():direction*(reach/length);
        }
        internal static MeasurePoint Tangent(MeasurementItem item,int index)
        {return item.Tangents!=null&&item.Tangents.Count==item.Points.Count?item.Tangents[index]:AutoTangent(item.Points,index);}
        internal static void EnsureTangents(MeasurementItem item)
        {if(item.Tangents==null||item.Tangents.Count!=item.Points.Count)item.Tangents=Enumerable.Range(0,item.Points.Count).Select(i=>AutoTangent(item.Points,i)).ToList();}
        internal static List<MeasurePoint> Sample(MeasurementItem item)
        {
            if(item.Points.Count<3||item.Points.Count>512)throw new InvalidOperationException("n점원은 테두리를 따라 3~512개 점을 지정하세요.");
            var result=new List<MeasurePoint>{item.Points[0]};
            for(int i=0;i<item.Points.Count;i++)
            {
                int j=(i+1)%item.Points.Count;MeasurePoint a=item.Points[i],b=item.Points[j];
                if(MeasurementGeometry.Distance(a,b)<.000001)throw new InvalidOperationException("이웃한 점이 겹칩니다. 점을 다른 위치로 옮기세요.");
                Subdivide(result,a,a+Tangent(item,i),b-Tangent(item,j),b,0);
            }
            result.RemoveAt(result.Count-1);return result;
        }
        private static void Subdivide(List<MeasurePoint> points,MeasurePoint a,MeasurePoint c,MeasurePoint d,MeasurePoint b,int depth)
        {
            // Bound polyline error in image coordinates, independent of zoom.
            double excess=MeasurementGeometry.Distance(a,c)+MeasurementGeometry.Distance(c,d)+MeasurementGeometry.Distance(d,b)-MeasurementGeometry.Distance(a,b);
            if(depth>=14||(excess<.002&&Math.Max(MeasurementGeometry.SegmentDistance(c,a,b),MeasurementGeometry.SegmentDistance(d,a,b))<.025))
            {points.Add(b);if(points.Count>20000)throw new InvalidOperationException("곡선이 너무 복잡합니다. 점 수나 곡률을 줄여 주세요.");return;}
            MeasurePoint ac=(a+c)*.5,cd=(c+d)*.5,db=(d+b)*.5,left=(ac+cd)*.5,right=(cd+db)*.5,middle=(left+right)*.5;
            Subdivide(points,a,ac,left,middle,depth+1);Subdivide(points,middle,right,db,b,depth+1);
        }
    }
}

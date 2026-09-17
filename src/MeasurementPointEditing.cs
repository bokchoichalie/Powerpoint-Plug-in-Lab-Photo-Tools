using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace LabPhotoTools
{
    internal sealed partial class MeasurementCanvas
    {
        private int editItemId,editPoint=-1,editHandle;
        private string pointDragBefore;
        internal void BeginPointEdit()
        {
            if(Selected.Count!=1){Tell("측정 도형이나 측정값을 하나 선택한 뒤 점편집을 누르세요.");return;}
            MeasurementItem item=Document.Items.Find(i=>Selected.Contains(i.Id));
            if(item==null||item.Kind=="circle_distance"){Tell("원 사이 거리는 원본 원을 선택해 편집하세요.");return;}
            SetTool("editpoints",false);editItemId=item.Id;editPoint=-1;editHandle=0;
            Tell("사각 점을 끌어 이동하세요. n점원은 주황 핸들을 끌어 곡률을 조절합니다. Enter: 편집 완료 · Esc: 이동 취소");Notify();
        }
        private MeasurementItem EditingItem()
        {return Selected.Count==1?Document.Items.Find(i=>Selected.Contains(i.Id)):null;}
        private void StartPointDrag(MeasurePoint point)
        {
            MeasurementItem item=EditingItem();if(item==null)return;
            int node=-1,handle=0;double best=9/ViewScale;
            for(int i=0;i<item.Points.Count;i++)
            {
                double d=MeasurementGeometry.Distance(point,item.Points[i]);if(d<=best){best=d;node=i;}
            }
            // Anchor centers win over their own collapsed (corner) handles.
            if(node<0&&item.Kind=="ncurve")for(int i=0;i<item.Points.Count;i++)foreach(int sign in new[]{-1,1})
            {
                double d=MeasurementGeometry.Distance(point,item.Points[i]+MeasurementSpline.Tangent(item,i)*sign);
                if(d<=best){best=d;node=i;handle=sign;}
            }
            if(node<0)
            {
                MeasurementItem found=Hit(point,false);
                if(found!=null&&found.Kind!="circle_distance"){Selected.Clear();Selected.Add(found.Id);editItemId=found.Id;editPoint=-1;Notify();}
                return;
            }
            editItemId=item.Id;editPoint=node;editHandle=handle;pointDragBefore=Document.Serialize();Capture=true;Invalidate();
        }
        private void MoveEditedPoint(MeasurePoint point)
        {
            MeasurementItem item=Document.Items.Find(i=>i.Id==editItemId);if(item==null||editPoint<0||editPoint>=item.Points.Count)return;
            List<MeasurePoint> oldPoints=item.Points,oldTangents=item.Tangents;
            item.Points=oldPoints.Select(p=>new MeasurePoint(p.X,p.Y)).ToList();
            item.Tangents=oldTangents.Select(p=>new MeasurePoint(p.X,p.Y)).ToList();
            try
            {
                if(editHandle!=0)
                {
                    MeasurementSpline.EnsureTangents(item);MeasurePoint vector=(point-item.Points[editPoint])*editHandle;
                    double limit=Math.Max(Document.Width,Document.Height)*2,length=MeasurementGeometry.Distance(vector,new MeasurePoint());
                    item.Tangents[editPoint]=length>limit?vector*(limit/length):vector;
                }
                else
                {
                    // Freeze the current handles first, then move them with their anchor.
                    if(item.Kind=="ncurve")MeasurementSpline.EnsureTangents(item);
                    item.Points[editPoint]=new MeasurePoint(Math.Max(0,Math.Min(Document.Width,point.X)),Math.Max(0,Math.Min(Document.Height,point.Y)));
                }
                foreach(MeasurementItem candidate in Document.Items)MeasurementGeometry.Build(candidate,Document);
                Notify();
            }
            catch(InvalidOperationException){item.Points=oldPoints;item.Tangents=oldTangents;Tell("이 위치에서는 도형을 계산할 수 없습니다. 점을 다른 위치로 옮기세요.");Invalidate();}
        }
        private void EndPointDrag(bool cancel)
        {
            if(pointDragBefore==null)return;string before=pointDragBefore;pointDragBefore=null;
            if(cancel)Document=MeasurementDocument.Deserialize(before);
            else if(Document.Serialize()!=before){undo.Add(before);if(undo.Count>100)undo.RemoveAt(0);redo.Clear();}
            Capture=false;Notify();
        }
        internal void SetPointCurvature(bool corner)
        {
            MeasurementItem item=EditingItem();
            if(Tool!="editpoints"||item==null||item.Id!=editItemId||item.Kind!="ncurve"||editPoint<0||editPoint>=item.Points.Count)
            {Tell("n점원을 점편집한 뒤 바꿀 점을 클릭하세요.");return;}
            string before=Document.Serialize();MeasurementSpline.EnsureTangents(item);
            MeasurePoint old=item.Tangents[editPoint];item.Tangents[editPoint]=corner?new MeasurePoint():MeasurementSpline.AutoTangent(item.Points,editPoint);
            try{MeasurementGeometry.Build(item,Document);undo.Add(before);if(undo.Count>100)undo.RemoveAt(0);redo.Clear();Notify();Focus();}
            catch(InvalidOperationException){item.Tangents[editPoint]=old;Tell("해당 곡률로 도형을 계산할 수 없습니다.");}
        }
        private void DrawEditPoints(Graphics g)
        {
            MeasurementItem item=EditingItem();if(item==null||item.Kind=="circle_distance")return;
            using(Pen line=new Pen(Color.Orange,1))using(Pen border=new Pen(Color.Black,1))
            for(int i=0;i<item.Points.Count;i++)
            {
                PointF anchor=ToScreen(item.Points[i]);
                if(item.Kind=="ncurve")foreach(int sign in new[]{-1,1})
                {
                    PointF handle=ToScreen(item.Points[i]+MeasurementSpline.Tangent(item,i)*sign);g.DrawLine(line,anchor,handle);
                    g.FillEllipse(Brushes.Orange,handle.X-4,handle.Y-4,8,8);g.DrawEllipse(border,handle.X-4,handle.Y-4,8,8);
                }
                g.FillRectangle(item.Id==editItemId&&i==editPoint?Brushes.Orange:Brushes.Cyan,anchor.X-4,anchor.Y-4,8,8);g.DrawRectangle(border,anchor.X-4,anchor.Y-4,8,8);
            }
        }
    }
}

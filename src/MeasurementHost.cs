using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace LabPhotoTools
{
    internal sealed class MeasurementSession : IDisposable
    {
        internal Bitmap Image;
        internal MeasurementDocument Document;
        internal double Rotation;
        public void Dispose(){if(Image!=null){Image.Dispose();Image=null;}}
    }
    public partial class PowerPointHost
    {
        private const string MeasurementPrefix="LABPHOTO_MEASURE_";
        public SelectionSnapshot ReadMeasurementSelection()
        {
            SelectionSnapshot snapshot=CurrentSlide();dynamic selected=app.ActiveWindow.Selection;
            if((int)selected.Type!=2)throw new InvalidOperationException("치수를 측정할 사진 한 장을 선택하세요.");
            dynamic range=(bool)selected.HasChildShapeRange?selected.ChildShapeRange:selected.ShapeRange;
            for(int i=1;i<=(int)range.Count;i++)CollectSelectedPhotos(range.Item(i),snapshot.Photos);
            if(snapshot.Photos.Count!=1)throw new InvalidOperationException("치수측정은 사진 한 장씩 가능합니다. 사진 한 장 또는 사진 한 장이 들어 있는 그룹을 선택하세요.");
            return snapshot;
        }
        private static object FindMeasurementPhoto(dynamic shapes,string name,int id)
        {
            object byName=null;
            for(int i=1;i<=(int)shapes.Count;i++)
            {
                dynamic s=shapes.Item(i);
                if((int)s.Type==6){object nested=FindMeasurementPhoto(s.GroupItems,name,id);if(nested!=null)return nested;}
                else if(IsPhoto(s)&&(string)s.Name==name){if((int)s.Id==id)return (object)s;byName=s;}
            }
            return byName;
        }
        internal static object TopMeasurementParent(dynamic shape)
        {
            dynamic top=shape;
            for(int depth=0;depth<32;depth++){try{dynamic parent=top.ParentGroup;if(parent==null)break;top=parent;}catch{break;}}
            return (object)top;
        }
        private static void FlattenMeasurementGroup(dynamic shape,List<object> leaves)
        {
            if((int)shape.Type!=6){leaves.Add((object)shape);return;}
            dynamic range=shape.Ungroup();List<object> children=new List<object>();
            for(int i=1;i<=(int)range.Count;i++)children.Add((object)range.Item(i));
            foreach(object child in children)FlattenMeasurementGroup(child,leaves);
        }
        private static string MeasurementSignature(dynamic photo)
        {
            // Cropping and flipping change the calibration image. Rotation and
            // proportional slide resizing do not invalidate a stored scale.
            List<string> values=new List<string>();dynamic f=photo.PictureFormat;
            // PowerPoint reports CropLeft/Right/Top/Bottom in the original
            // picture's points; these stay fixed when its group is resized.
            foreach(double v in new[]{(double)f.CropLeft,(double)f.CropRight,(double)f.CropTop,(double)f.CropBottom})values.Add(v.ToString("0.00",CultureInfo.InvariantCulture));
            values.Add(((int)photo.HorizontalFlip).ToString());values.Add(((int)photo.VerticalFlip).ToString());
            return string.Join("|",values);
        }
        public static MeasurementDocument ReadMeasurementData(object shape)
        {
            dynamic photo=shape;int count;
            if(!int.TryParse(TagValue(photo,MeasurementPrefix+"COUNT"),out count)||count==0)return null;
            if(count<0||count>20000)throw new InvalidOperationException("사진의 측정 정보가 손상되었습니다.");
            StringBuilder hex=new StringBuilder();for(int i=0;i<count;i++)hex.Append(TagValue(photo,MeasurementPrefix+i.ToString("D5",CultureInfo.InvariantCulture)));
            if(hex.Length%2!=0)throw new InvalidOperationException("사진의 측정 정보가 손상되었습니다.");
            byte[] data=new byte[hex.Length/2];for(int i=0;i<data.Length;i++)data[i]=byte.Parse(hex.ToString(i*2,2),NumberStyles.HexNumber,CultureInfo.InvariantCulture);
            return MeasurementDocument.Deserialize(Encoding.UTF8.GetString(data));
        }
        private static void WriteMeasurementData(dynamic photo,MeasurementDocument doc)
        {
            int oldCount;int.TryParse(TagValue(photo,MeasurementPrefix+"COUNT"),out oldCount);
            byte[] bytes=Encoding.UTF8.GetBytes(doc.Serialize());StringBuilder hex=new StringBuilder(bytes.Length*2);
            foreach(byte b in bytes)hex.Append(b.ToString("X2",CultureInfo.InvariantCulture));
            int count=(hex.Length+199)/200;if(count>20000)throw new InvalidOperationException("측정 정보가 너무 큽니다. 측정 도형을 줄이세요.");
            for(int i=0;i<count;i++)photo.Tags.Add(MeasurementPrefix+i.ToString("D5",CultureInfo.InvariantCulture),hex.ToString(i*200,Math.Min(200,hex.Length-i*200)));
            for(int i=count;i<Math.Min(20000,oldCount);i++)photo.Tags.Delete(MeasurementPrefix+i.ToString("D5",CultureInfo.InvariantCulture));
            photo.Tags.Add(MeasurementPrefix+"COUNT",count.ToString(CultureInfo.InvariantCulture));
        }
        internal MeasurementSession CreateMeasurementSession(SelectionSnapshot selection)
        {
            if(selection.Photos.Count!=1)throw new InvalidOperationException("사진 한 장을 선택하세요.");VerifyUnchanged(selection);
            PhotoSnapshot source=selection.Photos[0];MeasurementDocument doc=ReadMeasurementData(source.Shape);
            string directory=Engine.NewJobDirectory();dynamic copy=null;Bitmap image=null;
            try
            {
                string path=Path.Combine(directory,"measurement-preview.pptx");((dynamic)selection.Presentation).SaveCopyAs(path,24);
                copy=app.Presentations.Open(path,-1,0,0);dynamic slide=copy.Slides.Item((int)((dynamic)selection.Slide).SlideIndex);
                dynamic picture=FindMeasurementPhoto(slide.Shapes,source.Name,source.Id);
                if(picture==null)throw new InvalidOperationException("선택한 사진을 복사본에서 찾지 못했습니다.");
                List<object> leaves=new List<object>();FlattenMeasurementGroup(TopMeasurementParent(picture),leaves);
                picture=leaves.Cast<dynamic>().First(s=>IsPhoto(s)&&(string)s.Name==source.Name);
                // Keep canonical measurement coordinates, but display them with
                // the full slide rotation (including rotation inherited from groups).
                double rotation=(double)picture.Rotation;picture.Rotation=0f;
                string signature=MeasurementSignature(picture);double aspect=(double)picture.Width/(double)picture.Height;
                if(doc!=null&&(Math.Abs(aspect/(doc.Width/doc.Height)-1)>.01||doc.PhotoSignature!=signature))
                    throw new InvalidOperationException("스케일 설정 후 사진의 자르기·뒤집기·가로세로 비율이 바뀌었습니다. 사진을 원래 상태로 되돌리거나 측정 탭의 ‘측정 초기화’로 새로 보정하세요.");
                double factor=2400.0/Math.Max((double)picture.Width,(double)picture.Height);
                string png=Path.Combine(directory,"photo.png");
                picture.Export(png,2,(int)Math.Ceiling(selection.SlideWidth*factor),(int)Math.Ceiling(selection.SlideHeight*factor),1);
                using(Image loaded=System.Drawing.Image.FromFile(png))image=new Bitmap(loaded);
                // Shape.Export rounds its pixel dimensions independently.
                // Retain the exact local aspect ratio for geometric calculations.
                if(doc==null)doc=new MeasurementDocument {Width=image.Width,Height=image.Width/aspect,PhotoSignature=signature};
                return new MeasurementSession {Image=image,Document=doc,Rotation=rotation};
            }
            catch{if(image!=null)image.Dispose();throw;}
            finally{if(copy!=null){try{copy.Saved=-1;copy.Close();}catch{}}Engine.CleanJob(directory);}
        }
        public static MeasurePoint MeasurementToSlide(MeasurePoint p,MeasurementDocument doc,PhotoSnapshot photo)
        {
            double x=(p.X/doc.Width-.5)*photo.Width,y=(p.Y/doc.Height-.5)*photo.Height,t=photo.Rotation*Math.PI/180;
            return new MeasurePoint(photo.Left+photo.Width/2+x*Math.Cos(t)-y*Math.Sin(t),photo.Top+photo.Height/2+x*Math.Sin(t)+y*Math.Cos(t));
        }
        private static int OfficeRgb(int argb){Color c=Color.FromArgb(argb);return c.R|(c.G<<8)|(c.B<<16);}
        private static void MarkMeasurementShape(dynamic shape,string token,MeasurementItem item)
        {shape.Tags.Add("LABPHOTO_MEASUREMENT_OVERLAY",token);shape.Tags.Add("LABPHOTO_MEASUREMENT_ID",item.Id.ToString(CultureInfo.InvariantCulture));shape.Name="LabMeasure_"+item.Id+"_"+shape.Id;}
        private static HashSet<int> MeasurementShapeIds(dynamic slide)
        {HashSet<int> ids=new HashSet<int>();for(int i=1;i<=(int)slide.Shapes.Count;i++)ids.Add((int)slide.Shapes.Item(i).Id);return ids;}
        private static void RollbackMeasurementCopy(dynamic slide,HashSet<int> before)
        {for(int i=(int)slide.Shapes.Count;i>=1;i--)try{dynamic s=slide.Shapes.Item(i);if(!before.Contains((int)s.Id))s.Delete();}catch{}}
        private static object CopyMeasurementPhoto(PhotoSnapshot original,List<string> groupNames)
        {
            dynamic top=TopMeasurementParent(original.Shape);dynamic copy=top.Duplicate().Item(1);
            copy.Left=(float)((double)top.Left+18);copy.Top=(float)((double)top.Top+18);
            dynamic picture=(int)copy.Type==6?FindMeasurementPhoto(copy.GroupItems,original.Name,original.Id):(object)copy;
            if(picture==null)throw new InvalidOperationException("복사한 사진을 찾지 못했습니다.");
            // Give the target its own name before ungrouping. A surrounding
            // group may contain unrelated photos, shapes or text: discard those.
            string target="LabMeasuredPhoto_"+Guid.NewGuid().ToString("N");picture.Name=target;
            List<PhotoSnapshot> photos=new List<PhotoSnapshot>();CollectSelectedPhotos(copy,photos);
            bool keepNumber=photos.Count==1;
            List<object> leaves=new List<object>();FlattenMeasurementGroup(copy,leaves);
            foreach(dynamic leaf in leaves)
            {
                if((string)leaf.Name==target){picture=leaf;groupNames.Add(target);}
                else if(keepNumber&&!string.IsNullOrEmpty(TagValue(leaf,"LABPHOTO_NUMBER_STYLE")))
                {leaf.Name="LabNumber_copy_"+Guid.NewGuid().ToString("N");groupNames.Add((string)leaf.Name);}
                else leaf.Delete();
            }
            return (object)picture;
        }
        public object ApplyMeasurements(SelectionSnapshot selection,MeasurementDocument document)
        {
            if(selection.Photos.Count!=1||!document.HasScale)throw new InvalidOperationException("사진 한 장과 보정된 스케일이 필요합니다.");
            VerifyUnchanged(selection);MeasurementDocument doc=MeasurementDocument.Deserialize(document.Serialize());
            app.StartNewUndoEntry();dynamic slide=selection.Slide;HashSet<int> before=MeasurementShapeIds(slide);
            try
            {
                List<string> groupNames=new List<string>();
                dynamic picture=CopyMeasurementPhoto(selection.Photos[0],groupNames);
                PhotoSnapshot photo=SnapshotPhoto(picture);string token=Guid.NewGuid().ToString("N");
                // Exported pixels already include the photo's flip. Map the
                // rendered local picture axes through its current rotation.
                foreach(MeasurementItem item in doc.Items)
                {
                    MeasurementGeometryResult result=MeasurementGeometry.Build(item,doc);
                    foreach(MeasurementPath path in result.Paths)
                    {
                        if(path.Points.Count<2)continue;MeasurePoint first=MeasurementToSlide(path.Points[0],doc,photo);
                        dynamic builder=slide.Shapes.BuildFreeform(0,(float)first.X,(float)first.Y);
                        for(int i=1;i<path.Points.Count;i++){MeasurePoint p=MeasurementToSlide(path.Points[i],doc,photo);builder.AddNodes(0,0,(float)p.X,(float)p.Y);}
                        if(path.Closed)builder.AddNodes(0,0,(float)first.X,(float)first.Y);
                        dynamic shape=builder.ConvertToShape();MarkMeasurementShape(shape,token,item);shape.Fill.Visible=path.Closed&&item.Filled?-1:0;
                        if(path.Closed&&item.Filled){shape.Fill.Solid();shape.Fill.ForeColor.RGB=OfficeRgb(item.ColorArgb);shape.Fill.Transparency=.76f;}
                        shape.Line.Visible=-1;shape.Line.ForeColor.RGB=OfficeRgb(item.ColorArgb);shape.Line.Weight=(float)Math.Max(.15,item.LineWidth*photo.Width/800);
                        shape.Line.DashStyle=item.Dashed?4:1;groupNames.Add((string)shape.Name);
                    }
                    if(!string.IsNullOrEmpty(result.Text))
                    {
                        MeasurePoint p=MeasurementToSlide(result.Label,doc,photo);float textSize=(float)Math.Max(1,item.FontSize*photo.Width/800);
                        dynamic label=slide.Shapes.AddTextbox(1,(float)p.X+3,(float)p.Y+3,120f,24f);MarkMeasurementShape(label,token,item);
                        label.Line.Visible=0;label.Fill.Visible=-1;label.Fill.Solid();label.Fill.ForeColor.RGB=0x1E1814;label.Fill.Transparency=.3f;
                        dynamic tf=label.TextFrame2;tf.MarginLeft=3f;tf.MarginRight=3f;tf.MarginTop=2f;tf.MarginBottom=2f;tf.WordWrap=0;tf.AutoSize=1;
                        tf.TextRange.Text=(item.Kind=="text"?"":"#"+item.Id+" ")+result.Text;tf.TextRange.Font.Name="Arial";tf.TextRange.Font.Size=textSize;
                        tf.TextRange.Font.Fill.Visible=-1;tf.TextRange.Font.Fill.ForeColor.RGB=OfficeRgb(item.TextColorArgb);groupNames.Add((string)label.Name);
                    }
                }
                doc.PhotoSignature=MeasurementSignature(picture);WriteMeasurementData(picture,doc);
                dynamic saved=picture;if(groupNames.Count>1)
                {
                    dynamic group=slide.Shapes.Range(groupNames.Cast<object>().ToArray()).Group();group.Name="LabMeasurementGroup_"+group.Id;
                    group.Tags.Add("LABPHOTO_MEASUREMENT_GROUP","1");picture.Tags.Add("LABPHOTO_ATTACHED_MEASUREMENT","1");
                    saved=group;
                }
                app.ActiveWindow.View.GotoSlide((int)slide.SlideIndex);saved.Select(-1);return (object)saved;
            }
            catch{RollbackMeasurementCopy(slide,before);throw;}
        }
        public void ResetMeasurements(SelectionSnapshot selection)
        {
            // Keep the original and make a clean photo copy on the same slide.
            if(selection.Photos.Count!=1)throw new InvalidOperationException("사진 한 장을 선택하세요.");VerifyUnchanged(selection);
            dynamic slide=selection.Slide;HashSet<int> before=MeasurementShapeIds(slide);app.StartNewUndoEntry();
            try
            {
                List<string> keep=new List<string>();
                dynamic photo=CopyMeasurementPhoto(selection.Photos[0],keep);
                for(int i=(int)photo.Tags.Count;i>=1;i--){string name=(string)photo.Tags.Name(i);if(name.StartsWith(MeasurementPrefix,StringComparison.Ordinal)||name=="LABPHOTO_ATTACHED_MEASUREMENT")photo.Tags.Delete(name);}
                dynamic saved=photo;if(keep.Count>1){saved=slide.Shapes.Range(keep.Cast<object>().ToArray()).Group();saved.Name="LabNumberGroup_reset_"+saved.Id;}
                app.ActiveWindow.View.GotoSlide((int)slide.SlideIndex);saved.Select(-1);
            }
            catch{RollbackMeasurementCopy(slide,before);throw;}
        }
    }
}

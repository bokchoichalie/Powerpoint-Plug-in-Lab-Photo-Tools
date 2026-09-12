using System;
using System.Collections.Generic;
using System.Globalization;

namespace LabPhotoTools
{
    public static class NumberLabels
    {
        public static bool IsStyle(string style) { return style == "square" || style == "circle" || style == "paren" || style == "suffix"; }
        public static string Text(string style, int number)
        {
            if (number < 0) throw new ArgumentOutOfRangeException("number");
            string text = number.ToString(CultureInfo.InvariantCulture);
            return style == "paren" ? "(" + Alphabet(number) + ")" : style == "suffix" ? text + ")" : text;
        }
        public static string Alphabet(int zeroBasedNumber)
        {
            if (zeroBasedNumber < 0) throw new ArgumentOutOfRangeException("zeroBasedNumber");
            string result = string.Empty;
            int value = zeroBasedNumber;
            do
            {
                result = ((char)('A' + value % 26)).ToString() + result;
                value = value / 26 - 1;
            }
            while (value >= 0);
            return result;
        }
        public static string Caption(string style, int number)
        {
            return (style == "square" ? "□ " : style == "circle" ? "○ " : "") + Text(style, number);
        }
        public static int Next(IEnumerable<int> numbers)
        {
            int maximum = 10;
            foreach (int number in numbers) maximum = Math.Max(maximum, number);
            if (maximum == int.MaxValue) throw new InvalidOperationException("추가할 수 있는 번호 범위를 초과했습니다.");
            return maximum + 1;
        }
    }

    public partial class PowerPointHost
    {
        private SelectionSnapshot CurrentSlide()
        {
            try
            {
                dynamic slide = app.ActiveWindow.View.Slide;
                dynamic presentation = app.ActivePresentation;
                // SlideIndex rejects notes/master views before any mutation.
                int index = (int)slide.SlideIndex;
                if (index < 1) throw new InvalidOperationException();
                return new SelectionSnapshot { Slide = slide, Presentation = presentation,
                    SlideWidth = (double)presentation.PageSetup.SlideWidth,
                    SlideHeight = (double)presentation.PageSetup.SlideHeight };
            }
            catch (Exception ex) { throw new InvalidOperationException("편집할 슬라이드를 일반 편집 화면에서 연 뒤 다시 실행하세요.", ex); }
        }
        private static bool IsPhoto(dynamic shape)
        {
            int type = (int)shape.Type;
            if (type == 13 || type == 11) return true;
            if (type == 14)
            {
                int contained = (int)shape.PlaceholderFormat.ContainedType;
                return contained == 13 || contained == 11;
            }
            return false;
        }
        private static void CollectPhotos(dynamic shapes, List<PhotoSnapshot> photos)
        {
            for (int i = 1; i <= (int)shapes.Count; i++)
            {
                dynamic shape = shapes.Item(i);
                if ((int)shape.Type == 6) { CollectPhotos(shape.GroupItems, photos); continue; }
                if (!IsPhoto(shape)) continue;
                photos.Add(new PhotoSnapshot { Shape = shape, Id = (int)shape.Id, Index = (int)shape.ZOrderPosition,
                    Left = (double)shape.Left, Top = (double)shape.Top, Width = (double)shape.Width, Height = (double)shape.Height,
                    Rotation = (double)shape.Rotation, Name = (string)shape.Name, AlternativeText = (string)shape.AlternativeText });
            }
        }
        public SelectionSnapshot ReadAllPhotos()
        {
            SelectionSnapshot snapshot = CurrentSlide();
            CollectPhotos(((dynamic)snapshot.Slide).Shapes, snapshot.Photos);
            if (snapshot.Photos.Count == 0) throw new InvalidOperationException("현재 슬라이드에 배열할 사진이 없습니다.");
            if (snapshot.Photos.Count > 1000) throw new InvalidOperationException("한 번에 사진 1000장까지 배열할 수 있습니다.");
            return snapshot;
        }
        public void ArrangeAllPhotos()
        {
            SelectionSnapshot snapshot = ReadSelectedPhotosOrNull();
            if (snapshot == null) snapshot = ReadAllPhotos();
            if (snapshot.Photos.Count == 0)
                throw new InvalidOperationException("선택한 항목 중 배열할 사진이 없습니다. 사진을 선택하거나 선택을 해제한 뒤 다시 실행하세요.");
            List<PhotoBox> envelopes = new List<PhotoBox>();
            foreach (PhotoSnapshot photo in snapshot.Photos)
            {
                double radians = photo.Rotation * Math.PI / 180;
                double width = Math.Abs(photo.Width * Math.Cos(radians)) + Math.Abs(photo.Height * Math.Sin(radians));
                double height = Math.Abs(photo.Width * Math.Sin(radians)) + Math.Abs(photo.Height * Math.Cos(radians));
                envelopes.Add(new PhotoBox { Id = photo.Id, Left = photo.Left + (photo.Width - width) / 2,
                    Top = photo.Top + (photo.Height - height) / 2, Width = width, Height = height });
            }
            List<PhotoBox> plan = Layout.SmartArrange(envelopes, snapshot.SlideWidth, snapshot.SlideHeight);
            VerifyUnchanged(snapshot);
            app.StartNewUndoEntry();
            try
            {
                foreach (PhotoBox slot in plan)
                {
                    PhotoSnapshot photo = snapshot.Photos.Find(p => p.Id == slot.Id);
                    PhotoBox envelope = envelopes.Find(p => p.Id == slot.Id);
                    double scale = slot.Width / envelope.Width;
                    double width = photo.Width * scale, height = photo.Height * scale;
                    PhotoBox target = new PhotoBox { Id = photo.Id, Left = slot.Left + (slot.Width - width) / 2,
                        Top = slot.Top + (slot.Height - height) / 2, Width = width, Height = height };
                    if (!SetAttachedLabelGroupBox(photo, target)) SetBox(photo.Shape, target);
                }
            }
            catch
            {
                foreach (PhotoSnapshot photo in snapshot.Photos) { try { if (!SetAttachedLabelGroupBox(photo, photo.Box())) SetBox(photo.Shape, photo.Box()); } catch { } }
                throw;
            }
        }
        private static void CollectNumberShapes(dynamic shapes, List<object> labels)
        {
            for (int i = 1; i <= (int)shapes.Count; i++)
            {
                dynamic shape = shapes.Item(i);
                if ((int)shape.Type == 6) CollectNumberShapes(shape.GroupItems, labels);
                if (!string.IsNullOrEmpty((string)shape.Tags.Item("LABPHOTO_NUMBER_STYLE"))) labels.Add((object)shape);
            }
        }
        private static string TagValue(dynamic shape, string name)
        {
            try { return Convert.ToString(shape.Tags.Item(name), CultureInfo.InvariantCulture) ?? string.Empty; }
            catch { return string.Empty; }
        }
        private sealed class NumberPhotoCandidate
        {
            public PhotoSnapshot Photo;
            public object OrphanedLabelGroup;
        }
        private static void CollectUnlabelledPhotoCandidates(dynamic shapes, List<NumberPhotoCandidate> candidates)
        {
            for (int i = 1; i <= (int)shapes.Count; i++)
            {
                dynamic shape = shapes.Item(i);
                if ((int)shape.Type == 6)
                {
                    // A label group normally has two children. If the user
                    // deletes only its label, PowerPoint can leave a one-item
                    // group containing the tagged photograph. Treat that
                    // photograph as unlabelled based on the current contents.
                    if ((int)shape.GroupItems.Count == 1)
                    {
                        dynamic child = shape.GroupItems.Item(1);
                        if (IsPhoto(child) && !string.IsNullOrEmpty(TagValue(child, "LABPHOTO_ATTACHED_LABEL")))
                            candidates.Add(new NumberPhotoCandidate { Photo = SnapshotPhoto(child), OrphanedLabelGroup = (object)shape });
                    }
                    continue;
                }
                if (!IsPhoto(shape)) continue;
                // A top-level photograph cannot currently be attached to a
                // label group. Ignore any stale tag left after ungroup/delete.
                candidates.Add(new NumberPhotoCandidate { Photo = SnapshotPhoto(shape) });
            }
        }
        private static NumberPhotoCandidate NextUnlabelledPhoto(dynamic slide)
        {
            List<NumberPhotoCandidate> candidates = new List<NumberPhotoCandidate>();
            CollectUnlabelledPhotoCandidates(slide.Shapes, candidates);
            if (candidates.Count == 0) return null;
            candidates.Sort(delegate(NumberPhotoCandidate first, NumberPhotoCandidate second)
            {
                int top = first.Photo.Top.CompareTo(second.Photo.Top);
                return top != 0 ? top : first.Photo.Left.CompareTo(second.Photo.Left);
            });
            // Treat small vertical differences as one visual row. This keeps
            // a hand-arranged row ordered left-to-right even when it is not
            // perfectly level.
            List<List<NumberPhotoCandidate>> rows = new List<List<NumberPhotoCandidate>>();
            foreach (NumberPhotoCandidate candidate in candidates)
            {
                if (rows.Count == 0) { rows.Add(new List<NumberPhotoCandidate> { candidate }); continue; }
                List<NumberPhotoCandidate> row = rows[rows.Count - 1];
                double top = row[0].Photo.Top;
                double tolerance = Math.Max(8, Math.Min(row[0].Photo.Height, candidate.Photo.Height) * .30);
                if (candidate.Photo.Top - top <= tolerance) row.Add(candidate);
                else rows.Add(new List<NumberPhotoCandidate> { candidate });
            }
            foreach (List<NumberPhotoCandidate> row in rows)
                row.Sort(delegate(NumberPhotoCandidate first, NumberPhotoCandidate second) { return first.Photo.Left.CompareTo(second.Photo.Left); });
            return rows[0][0];
        }
        private static PhotoSnapshot PrepareUnlabelledPhoto(NumberPhotoCandidate candidate)
        {
            dynamic picture = candidate.Photo.Shape;
            if (candidate.OrphanedLabelGroup != null)
            {
                dynamic ungrouped = ((dynamic)candidate.OrphanedLabelGroup).Ungroup();
                picture = null;
                for (int i = 1; i <= (int)ungrouped.Count; i++)
                {
                    dynamic item = ungrouped.Item(i);
                    if ((int)item.Id == candidate.Photo.Id) { picture = item; break; }
                }
                if (picture == null)
                    throw new InvalidOperationException("라벨을 지운 사진을 다시 찾지 못했습니다. 사진을 선택 해제한 뒤 다시 시도하세요.");
            }
            try { picture.Tags.Delete("LABPHOTO_ATTACHED_LABEL"); } catch { }
            return SnapshotPhoto(picture);
        }
        private static void PlaceStandaloneLabel(dynamic shape, List<object> labels, float side, double slideWidth, double slideHeight)
        {
            float step = side + 6;
            int columns = Math.Max(1, (int)Math.Floor((slideWidth * .9 + 6) / step));
            int rows = Math.Max(1, (int)Math.Floor((slideHeight * .9 + 6) / step));
            for (int attempt = 0; attempt < columns * rows; attempt++)
            {
                float left = (float)(slideWidth * .05) + (attempt % columns) * step;
                float top = (float)(slideHeight * .05) + (attempt / columns) * step;
                bool occupied = false;
                foreach (dynamic label in labels)
                {
                    double right = left + side, bottom = top + side;
                    if (left < (double)label.Left + (double)label.Width && (double)label.Left < right &&
                        top < (double)label.Top + (double)label.Height && (double)label.Top < bottom)
                    { occupied = true; break; }
                }
                if (!occupied) { shape.Left = left; shape.Top = top; return; }
            }
            throw new InvalidOperationException("슬라이드에 새 번호 라벨을 놓을 공간이 없습니다. 기존 라벨을 정리하거나 새 슬라이드를 사용하세요.");
        }
        // PowerPoint scales every child of a group together. Scale the group
        // using the photograph's target scale, then translate it until that
        // photograph reaches its planned box. This preserves the label's
        // top-left attachment and its proportional size.
        private static bool SetAttachedLabelGroupBox(PhotoSnapshot photo, PhotoBox target)
        {
            dynamic picture = photo.Shape;
            if (string.IsNullOrEmpty(TagValue(picture, "LABPHOTO_ATTACHED_LABEL"))) return false;
            dynamic group;
            try { group = picture.ParentGroup; }
            catch { return false; }
            if (group == null || (int)group.Type != 6) return false;
            bool hasLabel = false;
            for (int i = 1; i <= (int)group.GroupItems.Count; i++)
                if (!string.IsNullOrEmpty(TagValue(group.GroupItems.Item(i), "LABPHOTO_NUMBER_STYLE"))) { hasLabel = true; break; }
            if (!hasLabel) return false;
            double currentWidth = (double)picture.Width;
            double currentHeight = (double)picture.Height;
            if (currentWidth <= 0 || currentHeight <= 0)
                throw new InvalidOperationException("라벨이 붙은 사진의 크기를 확인할 수 없습니다.");
            double widthScale = target.Width / currentWidth;
            double heightScale = target.Height / currentHeight;
            if (Math.Abs(widthScale - heightScale) > .01)
                throw new InvalidOperationException("라벨이 붙은 사진의 가로세로 비율이 변경되어 배열할 수 없습니다.");
            int lockAspectRatio = (int)group.LockAspectRatio;
            try
            {
                group.LockAspectRatio = 0;
                group.Width = (float)((double)group.Width * widthScale);
                group.Height = (float)((double)group.Height * heightScale);
                group.Left = (float)((double)group.Left + target.Left - (double)picture.Left);
                group.Top = (float)((double)group.Top + target.Top - (double)picture.Top);
            }
            finally { group.LockAspectRatio = lockAspectRatio; }
            return true;
        }
        public object AddNumberLabel(string style, int? requestedNumber)
        {
            return AddNumberLabel(style, requestedNumber, NumberLabelPreferences.Load());
        }
        public object AddNumberLabel(string style, int? requestedNumber, NumberLabelSettings settings)
        {
            if (!NumberLabels.IsStyle(style) || requestedNumber < 0) throw new ArgumentException("올바른 번호 형식을 선택하세요.");
            if (settings == null) throw new ArgumentNullException("settings");
            settings.Validate();
            SelectionSnapshot snapshot = CurrentSlide();
            dynamic slide = snapshot.Slide;
            List<object> labels = new List<object>();
            CollectNumberShapes(slide.Shapes, labels);
            List<int> numbers = new List<int>();
            foreach (dynamic label in labels)
            {
                if (!string.Equals(TagValue(label, "LABPHOTO_NUMBER_STYLE"), style, StringComparison.OrdinalIgnoreCase)) continue;
                int existing;
                if (int.TryParse(TagValue(label, "LABPHOTO_NUMBER"), out existing)) numbers.Add(existing);
            }
            int number = requestedNumber ?? NumberLabels.Next(numbers);
            NumberPhotoCandidate candidate = NextUnlabelledPhoto(slide);
            app.StartNewUndoEntry();
            dynamic shape = null;
            try
            {
                bool framed = style == "square" || style == "circle";
                shape = framed ? slide.Shapes.AddShape(style == "circle" ? 9 : 1, 0f, 0f, 48f, 48f)
                               : slide.Shapes.AddTextbox(1, 0f, 0f, 48f, 48f);
                shape.Name = "LabNumber_" + style + "_" + number + "_" + shape.Id;
                shape.Tags.Add("LABPHOTO_NUMBER_STYLE", style);
                shape.Tags.Add("LABPHOTO_NUMBER", number.ToString(CultureInfo.InvariantCulture));
                shape.Fill.Visible = framed ? -1 : 0;
                if (framed) { shape.Fill.Solid(); shape.Fill.ForeColor.RGB = 0xFFFFFF; shape.Fill.Transparency = 0f; }
                shape.Line.Visible = framed ? -1 : 0;
                if (framed) { shape.Line.ForeColor.RGB = 0; shape.Line.Weight = 1f; }
                dynamic frame = shape.TextFrame2;
                frame.MarginLeft = settings.Padding; frame.MarginRight = settings.Padding;
                frame.MarginTop = settings.Padding; frame.MarginBottom = settings.Padding;
                frame.VerticalAnchor = 3; frame.WordWrap = 0; frame.AutoSize = 0;
                frame.TextRange.Text = NumberLabels.Text(style, number);
                frame.TextRange.Font.Name = settings.FontName;
                frame.TextRange.Font.NameAscii = settings.FontName;
                frame.TextRange.Font.NameFarEast = settings.FontName;
                frame.TextRange.Font.Size = settings.FontSize;
                frame.TextRange.Font.Bold = -1;
                frame.TextRange.Font.Fill.Visible = -1;
                frame.TextRange.Font.Fill.Solid();
                frame.TextRange.Font.Fill.ForeColor.RGB = settings.OfficeColor;
                frame.TextRange.Font.Fill.Transparency = 0f;
                frame.TextRange.ParagraphFormat.Alignment = 2;
                // Keep the requested point size: measure actual Office text,
                // grow the square container, and leave AutoFit disabled.
                float side = settings.SquareSide((double)frame.TextRange.BoundWidth, (double)frame.TextRange.BoundHeight);
                if (side > snapshot.SlideWidth * .9 || side > snapshot.SlideHeight * .9)
                    throw new InvalidOperationException("이 글자 크기의 번호 라벨이 슬라이드보다 큽니다. 빠른 번호 매기기에서 글자 크기를 줄여 주세요.");
                shape.LockAspectRatio = 0;
                shape.Width = side; shape.Height = side;
                shape.LockAspectRatio = -1;
                shape.ZOrder(0);
                if (candidate == null)
                {
                    // A label is also useful as a standalone annotation. When
                    // no unlabelled photo remains, place it in the first free
                    // grid position exactly as earlier versions did.
                    PlaceStandaloneLabel(shape, labels, side, snapshot.SlideWidth, snapshot.SlideHeight);
                    shape.Select(-1);
                }
                else
                {
                    PhotoSnapshot photo = PrepareUnlabelledPhoto(candidate);
                    shape.Left = (float)photo.Left; shape.Top = (float)photo.Top;
                    dynamic group = slide.Shapes.Range(new object[] { photo.Name, (string)shape.Name }).Group();
                    ((dynamic)photo.Shape).Tags.Add("LABPHOTO_ATTACHED_LABEL", "1");
                    group.Name = "LabNumberGroup_" + style + "_" + number + "_" + group.Id;
                    group.Select(-1);
                }
                return (object)shape;
            }
            catch { if (shape != null) { try { shape.Delete(); } catch { } } throw; }
        }
    }
}


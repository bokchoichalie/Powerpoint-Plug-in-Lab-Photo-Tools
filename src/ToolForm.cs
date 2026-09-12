using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LabPhotoTools
{
    public sealed class ToolForm : Form
    {
        private const double PointsPerMm = 72.0 / 25.4;
        private readonly PowerPointHost host;
        private readonly SelectionSnapshot selection;
        private readonly string mode;
        private readonly Panel viewport = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Margin = Padding.Empty };
        private readonly TableLayoutPanel sheet = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Margin = Padding.Empty };
        private readonly TableLayoutPanel body = Stack();
        private readonly TableLayoutPanel previewPane = Stack();
        private readonly Label status = new Label { AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 8, 0, 8) };
        private readonly Button apply = new Button();
        private readonly Button close = new Button();
        private readonly PhotoPreview picture = new PhotoPreview { Dock = DockStyle.Fill };
        private readonly List<Control> wrapping = new List<Control>();
        private NumericUpDown angle, gridColumns, gridRows, columns, width, gapX, gapY;
        private CheckBox grid, uniform, horizontal, vertical;
        private TrackBar rotationSlider;
        private ImagePreviewSession previewSession;
        private Task backgroundTask;
        private CancellationTokenSource cancellation;
        private bool busy, closed, reflowing, closeRequested;
        private bool IsImage { get { return mode == "image" || mode == "background"; } }
        private float UiScale { get { return Math.Max(0.75f, Font.SizeInPoints / 10f * DeviceDpi / 96f); } }

        public ToolForm(PowerPointHost host, SelectionSnapshot selection, string mode)
        {
            this.host = host; this.selection = selection; this.mode = mode;
            if (mode != "image" && mode != "background" && mode != "layout" && mode != "spacing") throw new ArgumentException("Unknown tool mode.");
            SuspendLayout();
            AutoScaleDimensions = new SizeF(96, 96);
            AutoScaleMode = AutoScaleMode.Dpi;
            Font = new Font("맑은 고딕", 10f);
            Text = "Lab Photo Tools · " + (mode == "image" ? "사진 회전" : mode == "background" ? "배경지우기" : mode == "layout" ? "자동 배열" : "간격 조절");
            BackColor = Color.FromArgb(248, 249, 251);
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            ClientSize = IsImage ? new Size(1000, 690) : new Size(520, 680);
            MinimumSize = new Size(300, 260);
            TableLayoutPanel root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(14) };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Controls.Add(root);
            root.Controls.Add(viewport, 0, 0); viewport.Controls.Add(sheet);
            sheet.Controls.Add(body);
            AddText(mode == "image" ? "사진 회전" : mode == "background" ? "배경지우기" : mode == "layout" ? "사진을 가지런하게" : "사진 사이 여백 맞추기", true);
            AddText("선택한 사진 " + selection.Photos.Count + "장", false);
            if (mode == "image") BuildRotation();
            else if (mode == "background")
            {
                AddText("선택한 사진의 배경을 투명하게 만듭니다. 첫 사진의 결과를 미리 확인할 수 있습니다.", false);
                AddText("완료 후 ‘복사본에 적용’을 누르면 선택한 모든 사진의 배경을 지웁니다. 사진의 각도와 비율은 유지됩니다.", false);
                Button retry = new Button { Text = "미리보기 다시 시도", AutoSize = true, Padding = new Padding(8, 5, 8, 5) };
                retry.Click += async delegate { if (!busy && (backgroundTask == null || backgroundTask.IsCompleted)) await PrepareBackground(); };
                Add(retry);
            }
            else BuildLayout();
            if (IsImage)
            {
                Label title = TextLabel("첫 사진 · 자동 미리보기", true);
                previewPane.Controls.Add(title);
                previewPane.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                previewPane.Controls.Add(picture);
                previewPane.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
                Label caption = TextLabel(mode == "image" ? "격자는 미리보기에만 표시됩니다.\n선택한 모든 사진에 같은 회전 각도를 적용합니다." : "체크무늬는 투명 영역입니다.\n결과는 원본 슬라이드의 복사본에 적용합니다.", false);
                previewPane.Controls.Add(caption);
                previewPane.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                previewPane.RowCount = 3;
                previewPane.AutoSize = false;
                sheet.Controls.Add(previewPane);
            }
            root.Controls.Add(status, 0, 1);
            FlowLayoutPanel buttons = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, WrapContents = true, Margin = Padding.Empty };
            ConfigureButton(apply, IsImage ? "복사본에 적용" : "적용", true);
            ConfigureButton(close, "닫기", false);
            buttons.Controls.Add(apply); buttons.Controls.Add(close);
            root.Controls.Add(buttons, 0, 2);
            apply.Enabled = !IsImage;
            apply.Click += async delegate { await Apply(); };
            close.Click += delegate { Close(); };
            CancelButton = close;
            status.Text = IsImage ? "미리보기를 준비합니다. 결과는 원본 슬라이드의 복사본에 적용됩니다." : "슬라이드 밖으로 넘치는 설정은 적용되지 않습니다.";
            viewport.SizeChanged += delegate { Reflow(); };
            body.SizeChanged += delegate { Reflow(); };
            Shown += async delegate { FitToScreen(); Reflow(); if (IsImage) await LoadPreview(); };
            LocationChanged += delegate { if (IsHandleCreated && !reflowing) FitToScreen(false); };
            FormClosing += delegate(object sender, FormClosingEventArgs e)
            {
                if (!busy) return;
                closeRequested = true;
                if (cancellation != null) cancellation.Cancel();
                e.Cancel = true;
                status.Text = "작업을 취소하고 정리하는 중입니다…";
            };
            ResumeLayout(true);
            Reflow();
        }
        private static TableLayoutPanel Stack()
        {
            TableLayoutPanel panel = new TableLayoutPanel { ColumnCount = 1, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Dock = DockStyle.Fill, Margin = Padding.Empty, GrowStyle = TableLayoutPanelGrowStyle.AddRows };
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            return panel;
        }
        private void FitToScreen(bool center = true)
        {
            Rectangle work = Screen.FromHandle(Handle).WorkingArea;
            Size maximum = new Size(Math.Max(200, work.Width - 16), Math.Max(180, work.Height - 16));
            MinimumSize = new Size(Math.Min((int)(300 * UiScale), maximum.Width), Math.Min((int)(260 * UiScale), maximum.Height));
            if (Width > maximum.Width || Height > maximum.Height) Size = new Size(Math.Min(Width, maximum.Width), Math.Min(Height, maximum.Height));
            if (center) Location = new Point(work.Left + (work.Width - Width) / 2, work.Top + (work.Height - Height) / 2);
        }
        private void Reflow()
        {
            if (reflowing || IsDisposed || viewport.ClientSize.Width <= 0) return;
            reflowing = true;
            try
            {
                float scale = UiScale;
                bool wide = IsImage && viewport.ClientSize.Width >= 780 * scale;
                sheet.SuspendLayout();
                sheet.MinimumSize = new Size((int)(250 * scale), 0);
                sheet.ColumnCount = wide ? 2 : 1;
                sheet.RowCount = wide || !IsImage ? 1 : 2;
                sheet.ColumnStyles.Clear(); sheet.RowStyles.Clear();
                sheet.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, wide ? 44 : 100));
                if (wide) sheet.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 56));
                sheet.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                if (IsImage && !wide) sheet.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                sheet.SetCellPosition(body, new TableLayoutPanelCellPosition(0, 0));
                if (IsImage)
                {
                    sheet.SetCellPosition(previewPane, new TableLayoutPanelCellPosition(wide ? 1 : 0, wide ? 0 : 1));
                    previewPane.Margin = new Padding(wide ? (int)(16 * scale) : 0, wide ? 0 : (int)(14 * scale), 0, 0);
                    previewPane.MinimumSize = new Size(0, (int)(340 * scale));
                    previewPane.Height = (int)(wide ? 470 * scale : 340 * scale);
                }
                int contentWidth = Math.Max((int)(230 * scale), (int)((viewport.ClientSize.Width - SystemInformation.VerticalScrollBarWidth - 8) * (wide ? .43 : 1)));
                foreach (Control control in wrapping)
                    control.MaximumSize = new Size(control.Parent == body ? contentWidth : Math.Max(120, contentWidth - 8), 0);
                status.MaximumSize = new Size(Math.Max(80, viewport.ClientSize.Width), 0);
                sheet.ResumeLayout(true);
            }
            finally { reflowing = false; }
        }
        private Label TextLabel(string text, bool heading)
        {
            Label label = new Label { Text = text, AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 4, 0, 10), ForeColor = Color.FromArgb(41, 54, 69) };
            if (heading) label.Font = new Font(Font.FontFamily, Font.SizeInPoints * 1.35f, FontStyle.Bold);
            wrapping.Add(label);
            return label;
        }
        private void AddText(string text, bool heading) { Add(TextLabel(text, heading)); }
        private void Add(Control control)
        {
            int row = body.Controls.Count;
            body.RowCount = row + 1;
            body.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            body.Controls.Add(control, 0, row);
        }
        private CheckBox AddCheck(string text, bool value)
        {
            CheckBox check = new CheckBox { Text = text, Checked = value, AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 6, 0, 8) };
            wrapping.Add(check); Add(check); return check;
        }
        private NumericUpDown AddNumber(string label, decimal min, decimal max, decimal value, int decimals, decimal step)
        {
            TableLayoutPanel field = Stack();
            field.Margin = new Padding(0, 2, 0, 10);
            Label caption = TextLabel(label, false); caption.Margin = new Padding(0, 0, 0, 4);
            field.Controls.Add(caption, 0, 0); field.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            NumericUpDown number = new NumericUpDown { Minimum = min, Maximum = max, Value = value, DecimalPlaces = decimals, Increment = step, ThousandsSeparator = true, Dock = DockStyle.Top, AccessibleName = label, Margin = Padding.Empty };
            field.Controls.Add(number, 0, 1); field.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Add(field); return number;
        }
        private void BuildRotation()
        {
            angle = AddNumber("추가 회전 각도 (°)", -45, 45, 0, 1, .1m);
            rotationSlider = new TrackBar { Minimum = -450, Maximum = 450, SmallChange = 1, LargeChange = 10, TickFrequency = 150, Dock = DockStyle.Top, AutoSize = true, Margin = Padding.Empty, AccessibleName = "추가 회전 각도 슬라이더 (0.1도 단위)" };
            Add(rotationSlider);
            AddText("−45°                 0°                 +45°\n0.1° 단위 · +는 시계 방향입니다.\n원래 비율을 유지하며 빈 가장자리를 자릅니다.", false);
            grid = AddCheck("미리보기에 점선 격자 표시", true);
            gridColumns = AddNumber("가로 칸 수 (세로선)", 1, 50, 4, 0, 1);
            gridRows = AddNumber("세로 칸 수 (가로선)", 1, 50, 4, 0, 1);
            rotationSlider.ValueChanged += delegate { angle.Value = rotationSlider.Value / 10m; };
            angle.ValueChanged += delegate { rotationSlider.Value = (int)Math.Round(angle.Value * 10); RefreshPreview(); };
            grid.CheckedChanged += delegate { gridColumns.Enabled = gridRows.Enabled = grid.Checked; RefreshPreview(); };
            gridColumns.ValueChanged += delegate { RefreshPreview(); }; gridRows.ValueChanged += delegate { RefreshPreview(); };
        }
        private void BuildLayout()
        {
            if (mode == "layout")
            {
                columns = AddNumber("한 행의 사진 수 (열)", 1, 50, Math.Min(3, selection.Photos.Count), 0, 1);
                uniform = AddCheck("사진 너비 통일 · 세로 비율 유지", true);
                width = AddNumber("사진 너비 (mm)", 1, 1000, Math.Max(1m, Math.Min(45m, (decimal)(selection.Photos[0].Width / PointsPerMm))), 1, 1);
                uniform.CheckedChanged += delegate { width.Enabled = uniform.Checked; };
                AddText("현재 위치를 기준으로 위에서 아래, 왼쪽에서 오른쪽 순서로 배열합니다.", false);
            }
            else
            {
                horizontal = AddCheck("가로 간격 조절", true); vertical = AddCheck("행 사이 간격 조절", true);
                AddText("사진 크기를 유지합니다. 위쪽 위치가 비슷한 사진을 같은 행으로 판단합니다.", false);
            }
            gapX = AddNumber("사진 사이 가로 간격 (mm)", 0, 1000, 3, 1, .5m);
            gapY = AddNumber("행 사이 세로 간격 (mm)", 0, 1000, 3, 1, .5m);
            if (mode == "spacing")
            {
                horizontal.CheckedChanged += delegate { gapX.Enabled = horizontal.Checked; };
                vertical.CheckedChanged += delegate { gapY.Enabled = vertical.Checked; };
            }
        }
        private async Task LoadPreview()
        {
            status.Text = "첫 사진을 읽는 중입니다…";
            await Task.Yield();
            if (closed) return;
            try
            {
                previewSession = host.CreatePreviewSession(selection);
                RefreshPreview();
                if (mode == "background") await PrepareBackground();
                else { apply.Enabled = true; status.Text = "각도를 바꾸면 바로 미리봅니다. 결과는 원본 슬라이드의 복사본에 적용됩니다."; }
            }
            catch (Exception ex) { if (!closed) { status.Text = "미리보기를 읽지 못했습니다: " + ex.Message; apply.Enabled = true; } }
        }
        private void RefreshPreview()
        {
            if (closed || previewSession == null) return;
            picture.Source = mode == "background" && previewSession.BackgroundRemoved != null ? previewSession.BackgroundRemoved : previewSession.Original;
            picture.Angle = selection.Photos[0].Rotation + (mode == "image" ? (double)angle.Value : 0);
            picture.Crop = mode == "image";
            picture.ShowGrid = mode == "image" && grid.Checked;
            if (mode == "image") { picture.GridColumns = (int)gridColumns.Value; picture.GridRows = (int)gridRows.Value; }
            picture.Invalidate();
        }
        private async Task PrepareBackground()
        {
            if (previewSession == null) { await LoadPreview(); return; }
            if (closed) return;
            apply.Enabled = false; status.Text = "배경을 지우는 중입니다… 완료 전에는 원본이 보입니다.";
            try
            {
                backgroundTask = previewSession.PrepareBackgroundAsync();
                await backgroundTask;
                if (!closed) { RefreshPreview(); status.Text = "배경지우기 미리보기가 준비되었습니다. 복사본에 적용해 주세요."; }
            }
            catch (OperationCanceledException) { }
            catch (Exception ex) { if (!closed) status.Text = "미리보기 오류: " + ex.Message; }
            finally { if (!closed) apply.Enabled = true; }
        }
        private async Task Apply()
        {
            if (busy) return;
            busy = true; body.Enabled = apply.Enabled = false; close.Text = "취소";
            cancellation = new CancellationTokenSource();
            bool complete = false;
            try
            {
                status.Text = IsImage ? "이 PC에서 사진을 처리하는 중입니다…" : "사진 배치를 적용합니다…";
                if (IsImage) await host.ApplyImagesAsync(selection, mode == "background", mode == "image", mode == "image" ? (double)angle.Value : 0, cancellation.Token);
                else
                {
                    List<PhotoBox> plan = mode == "layout"
                        ? LabPhotoTools.Layout.Arrange(selection.Boxes(), (int)columns.Value, (double)gapX.Value * PointsPerMm, (double)gapY.Value * PointsPerMm, uniform.Checked, (double)width.Value * PointsPerMm, selection.SlideWidth, selection.SlideHeight)
                        : LabPhotoTools.Layout.Space(selection.Boxes(), (double)gapX.Value * PointsPerMm, (double)gapY.Value * PointsPerMm, horizontal.Checked, vertical.Checked, selection.SlideWidth, selection.SlideHeight);
                    host.ApplyLayout(selection, plan);
                }
                complete = true;
            }
            catch (OperationCanceledException) { status.Text = "취소했습니다."; }
            catch (Exception ex) { status.Text = ex.Message; }
            finally { busy = false; body.Enabled = apply.Enabled = true; close.Text = "닫기"; cancellation.Dispose(); cancellation = null; }
            if (complete || closeRequested) { DialogResult = complete ? DialogResult.OK : DialogResult.Cancel; Close(); }
        }
        private void ConfigureButton(Button button, string text, bool primary)
        {
            button.Text = text; button.AutoSize = true; button.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            button.Padding = new Padding(16, 7, 16, 7); button.Margin = new Padding(6, 0, 0, 0);
            button.FlatStyle = FlatStyle.Flat;
            button.BackColor = primary ? Color.FromArgb(25, 98, 180) : Color.White;
            button.ForeColor = primary ? Color.White : Color.FromArgb(35, 53, 73);
        }
        protected override void Dispose(bool disposing)
        {
            if (disposing && !closed)
            {
                closed = true; picture.Source = null;
                if (previewSession != null) previewSession.Dispose();
                if (cancellation != null) { cancellation.Cancel(); cancellation.Dispose(); }
            }
            base.Dispose(disposing);
        }
    }
}


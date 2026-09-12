using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Windows.Forms;

namespace LabPhotoTools
{
    internal sealed class RibbonIcons : AxHost
    {
        private static readonly Dictionary<string, object> cache = new Dictionary<string, object>();
        private static readonly Dictionary<int, object> colorCache = new Dictionary<int, object>();
        private RibbonIcons() : base("") { }
        internal static object ColorSwatch(Color color)
        {
            object value;
            if (!colorCache.TryGetValue(color.ToArgb(), out value))
            {
                using (Bitmap image = new Bitmap(24, 24, PixelFormat.Format32bppArgb))
                using (Graphics g = Graphics.FromImage(image))
                using (Brush brush = new SolidBrush(color))
                {
                    g.Clear(Color.Transparent); g.FillRectangle(brush, 2, 2, 19, 19); g.DrawRectangle(Pens.Gray, 2, 2, 19, 19);
                    value = GetIPictureDispFromPicture(image);
                }
                if (colorCache.Count > 96) colorCache.Clear();
                colorCache.Add(color.ToArgb(), value);
            }
            return value;
        }
        public static object Get(string id)
        {
            object value;
            if (!cache.TryGetValue(id, out value))
            {
                using (Bitmap image = Draw(id)) value = GetIPictureDispFromPicture(image);
                cache.Add(id, value);
            }
            return value;
        }
        internal static Bitmap Draw(string id)
        {
            if(id=="labPhotoMeasure")return MeasurementIcons.Draw("ruler",64);
            Bitmap bitmap = new Bitmap(64, 64, PixelFormat.Format32bppArgb);
            using (Graphics g = Graphics.FromImage(bitmap))
            using (Pen ink = new Pen(Color.FromArgb(40, 75, 110), 2.8f))
            using (SolidBrush blue = new SolidBrush(Color.FromArgb(53, 126, 201)))
            using (SolidBrush light = new SolidBrush(Color.FromArgb(229, 240, 252)))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);
                if (id == "labNumberSettings")
                {
                    using (Font font = new Font("Arial", 38, FontStyle.Bold, GraphicsUnit.Pixel))
                        g.DrawString("A", font, blue, 14, 0);
                    g.FillRectangle(Brushes.IndianRed, 6, 48, 14, 9);
                    g.FillRectangle(Brushes.SeaGreen, 25, 48, 14, 9);
                    g.FillRectangle(blue, 44, 48, 14, 9);
                }
                else if (id == "labPhotoGrid")
                {
                    for (int row = 0; row < 3; row++)
                        for (int col = 0; col < 3; col++)
                        {
                            Rectangle box = new Rectangle(5 + col * 19, 5 + row * 19, 14, 14);
                            g.FillRectangle(light, box); g.DrawRectangle(ink, box);
                        }
                }
                else if (id == "labPhotoSpacing")
                {
                    g.FillRectangle(light, 2, 15, 16, 34); g.DrawRectangle(ink, 2, 15, 16, 34);
                    g.FillRectangle(light, 46, 15, 16, 34); g.DrawRectangle(ink, 46, 15, 16, 34);
                    g.DrawLine(ink, 24, 32, 40, 32);
                    g.DrawLines(ink, new[] { new Point(28, 26), new Point(22, 32), new Point(28, 38) });
                    g.DrawLines(ink, new[] { new Point(36, 26), new Point(42, 32), new Point(36, 38) });
                }
                else if (id == "labPhotoMagic")
                {
                    using (Pen wand = new Pen(Color.FromArgb(65, 65, 101), 9)) g.DrawLine(wand, 12, 53, 40, 25);
                    using (Pen tip = new Pen(Color.FromArgb(166, 120, 231), 10)) g.DrawLine(tip, 37, 28, 45, 20);
                    Star(g, 19, 13, 7); Star(g, 50, 43, 7); Star(g, 50, 8, 5);
                }
                else if (id.StartsWith("number_", StringComparison.Ordinal))
                {
                    string[] parts = id.Split('_');
                    string style = parts[1];
                    int number = style == "paren" ? 0 : 1;
                    if (parts.Length > 2) int.TryParse(parts[2], out number);
                    if (style == "square") { g.FillRectangle(Brushes.White, 6, 6, 52, 52); g.DrawRectangle(ink, 6, 6, 52, 52); }
                    if (style == "circle") { g.FillEllipse(Brushes.White, 6, 6, 52, 52); g.DrawEllipse(ink, 6, 6, 52, 52); }
                    string text = NumberLabels.Text(style, number);
                    using (Font font = new Font("Arial", text.Length > 2 ? 20 : 25, FontStyle.Bold, GraphicsUnit.Pixel))
                    using (StringFormat format = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
                        g.DrawString(text, font, Brushes.Black, new RectangleF(1, 2, 62, 59), format);
                }
                else if (id == "labPhotoBackground")
                {
                    for (int y = 0; y < 4; y++) for (int x = 0; x < 4; x++)
                        g.FillRectangle((x + y) % 2 == 0 ? Brushes.White : Brushes.LightGray, 5 + x * 13, 5 + y * 13, 13, 13);
                    g.DrawRectangle(ink, 5, 5, 52, 52);
                    Point[] eraser = { new Point(25, 41), new Point(43, 17), new Point(58, 29), new Point(40, 53) };
                    g.FillPolygon(light, eraser); g.DrawPolygon(ink, eraser);
                    g.DrawLine(ink, 32, 31, 47, 43);
                }
                else
                {
                    g.TranslateTransform(31, 35); g.RotateTransform(-12);
                    g.FillRectangle(light, -18, -16, 36, 31); g.DrawRectangle(ink, -18, -16, 36, 31);
                    g.DrawLines(ink, new[] { new Point(-13, 9), new Point(-3, -2), new Point(5, 5), new Point(11, -2) });
                    g.ResetTransform(); g.DrawArc(ink, 11, 3, 42, 42, 195, 165);
                    g.FillPolygon(blue, new[] { new Point(46, 17), new Point(59, 23), new Point(54, 9) });
                }
            }
            return bitmap;
        }
        private static void Star(Graphics g, int x, int y, int r)
        {
            g.FillPolygon(Brushes.Goldenrod, new[] { new Point(x, y-r), new Point(x+2,y-2), new Point(x+r,y), new Point(x+2,y+2), new Point(x,y+r), new Point(x-2,y+2), new Point(x-r,y), new Point(x-2,y-2) });
        }
    }
}


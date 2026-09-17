using System;
using System.Drawing;
using System.IO;
using System.Text;
using System.Web.Script.Serialization;

namespace LabPhotoTools
{
    public sealed class NumberLabelSettings
    {
        // 17 pt type uses an 0.85 cm square/circle: 0.05 cm per point.
        public const float ContainerCentimetersPerPoint = .05f;
        public const float PointsPerCentimeter = 72f / 2.54f;
        public string FontName { get; set; }
        public float FontSize { get; set; }
        public int TextColorArgb { get; set; }
        public int BorderColorArgb { get; set; }
        public int FillColorArgb { get; set; }
        public float BorderWidth { get; set; }
        public string BorderMode { get; set; }
        public string FillMode { get; set; }

        public NumberLabelSettings()
        {
            FontName = "Arial";
            FontSize = 18f;
            TextColorArgb = Color.Black.ToArgb();
            BorderColorArgb=Color.Black.ToArgb();FillColorArgb=Color.White.ToArgb();BorderWidth=1;
            BorderMode=FillMode="default";
        }
        public NumberLabelSettings Copy()
        {
            return new NumberLabelSettings { FontName=FontName,FontSize=FontSize,TextColorArgb=TextColorArgb,
                BorderColorArgb=BorderColorArgb,FillColorArgb=FillColorArgb,BorderWidth=BorderWidth,BorderMode=BorderMode,FillMode=FillMode };
        }
        public void Validate()
        {
            if (string.IsNullOrWhiteSpace(FontName) || FontName.Length > 128 || FontName.IndexOf('\0') >= 0)
                throw new ArgumentException("글꼴을 선택해 주세요.");
            if (float.IsNaN(FontSize) || float.IsInfinity(FontSize) || FontSize < 1 || FontSize > 400)
                throw new ArgumentException("글자 크기는 1~400 pt 사이로 입력해 주세요.");
            if(float.IsNaN(BorderWidth)||float.IsInfinity(BorderWidth)||BorderWidth<0||BorderWidth>20)
                throw new ArgumentException("테두리 굵기는 0~20 pt 사이로 입력해 주세요. 0은 테두리 없음입니다.");
            if(!ValidMode(BorderMode)||!ValidMode(FillMode))throw new ArgumentException("테두리·바탕색 설정을 확인해 주세요.");
        }
        private static bool ValidMode(string mode){return mode=="default"||mode=="none"||mode=="color";}
        public static int ToOfficeColor(int argb){Color c=Color.FromArgb(argb);return c.R|(c.G<<8)|(c.B<<16);}
        public bool ShowBorder(bool framed){return BorderWidth>0&&(BorderMode=="color"||(BorderMode=="default"&&framed));}
        public bool ShowFill(bool framed){return FillMode=="color"||(FillMode=="default"&&framed);}
        public Color TextColor { get { return Color.FromArgb(255, Color.FromArgb(TextColorArgb)); } }
        public int OfficeColor { get { Color color = TextColor; return color.R | (color.G << 8) | (color.B << 16); } }
        public float BaseContainerSide { get { Validate(); return FontSize * ContainerCentimetersPerPoint * PointsPerCentimeter; } }
        public float Padding { get { return Math.Max(1, FontSize * .1f); } }
        public float SquareSide(double textWidth, double textHeight)
        {
            Validate();
            if (double.IsNaN(textWidth) || double.IsInfinity(textWidth) || double.IsNaN(textHeight) || double.IsInfinity(textHeight) || textWidth < 0 || textHeight < 0)
                throw new InvalidOperationException("번호 텍스트의 크기를 확인할 수 없습니다.");
            // Keep the chosen font-to-container ratio for ordinary labels. A
            // longer number grows only enough to keep all of its text visible.
            return (float)Math.Max(BaseContainerSide, Math.Max(textWidth, textHeight) + 2 * Padding);
        }
    }

    public static class NumberLabelPreferences
    {
        public static string SettingsPath
        {
            get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "LabPhotoTools", "preferences", "number-labels.json"); }
        }
        public static NumberLabelSettings Load() { return Load(SettingsPath); }
        public static NumberLabelSettings Load(string path)
        {
            try
            {
                if (!File.Exists(path) || new FileInfo(path).Length > 16384) return new NumberLabelSettings();
                NumberLabelSettings settings = new JavaScriptSerializer().Deserialize<NumberLabelSettings>(File.ReadAllText(path, Encoding.UTF8));
                if (settings == null) return new NumberLabelSettings();
                settings.Validate();
                return settings;
            }
            catch (IOException) { return new NumberLabelSettings(); }
            catch (UnauthorizedAccessException) { return new NumberLabelSettings(); }
            catch (ArgumentException) { return new NumberLabelSettings(); }
            catch (InvalidOperationException) { return new NumberLabelSettings(); }
        }
        public static void Save(NumberLabelSettings settings) { Save(SettingsPath, settings); }
        public static void Save(string path, NumberLabelSettings settings)
        {
            settings.Validate();
            string directory = Path.GetDirectoryName(Path.GetFullPath(path));
            Directory.CreateDirectory(directory);
            string temporary = Path.Combine(directory, "number-labels-" + Guid.NewGuid().ToString("N") + ".tmp");
            try
            {
                // Serialize the user preferences only; computed properties
                // such as Color/OfficeColor are not part of the saved format.
                string json = new JavaScriptSerializer().Serialize(new { FontName=settings.FontName,FontSize=settings.FontSize,TextColorArgb=settings.TextColorArgb,
                    BorderColorArgb=settings.BorderColorArgb,FillColorArgb=settings.FillColorArgb,BorderWidth=settings.BorderWidth,BorderMode=settings.BorderMode,FillMode=settings.FillMode });
                File.WriteAllText(temporary, json, new UTF8Encoding(false));
                if (File.Exists(path)) File.Replace(temporary, path, null);
                else File.Move(temporary, path);
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }
    }
}

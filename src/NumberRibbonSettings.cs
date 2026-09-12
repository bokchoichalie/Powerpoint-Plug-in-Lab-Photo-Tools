using System;
using System.Drawing;
using System.Globalization;
using System.Linq;

namespace LabPhotoTools
{
    internal static class NumberRibbonSettings
    {
        internal static readonly string[] Fonts = FontFamily.Families.Select(f => f.Name).Distinct().OrderBy(n => n, StringComparer.CurrentCultureIgnoreCase).ToArray();
        internal static readonly Color[] Colors = {
            Color.Black, Color.White, Color.FromArgb(68,68,68), Color.FromArgb(128,128,128), Color.FromArgb(192,192,192), Color.FromArgb(31,78,121), Color.FromArgb(0,112,192), Color.FromArgb(0,176,240),
            Color.FromArgb(192,0,0), Color.Red, Color.FromArgb(255,128,0), Color.FromArgb(255,192,0), Color.Yellow, Color.FromArgb(146,208,80), Color.FromArgb(0,176,80), Color.FromArgb(0,112,60),
            Color.FromArgb(112,48,160), Color.FromArgb(160,64,192), Color.FromArgb(255,0,128), Color.FromArgb(255,128,160), Color.FromArgb(0,128,128), Color.FromArgb(0,192,192), Color.FromArgb(139,69,19), Color.FromArgb(192,144,96),
            Color.FromArgb(222,235,247), Color.FromArgb(180,198,231), Color.FromArgb(226,239,218), Color.FromArgb(198,224,180), Color.FromArgb(255,230,153), Color.FromArgb(248,203,173), Color.FromArgb(244,176,132), Color.FromArgb(228,196,240)
        };
        internal static readonly string[] ColorNames = {
            "검정", "흰색", "진한 회색", "회색", "연한 회색", "남색", "파랑", "하늘색",
            "진한 빨강", "빨강", "주황", "금색", "노랑", "연두", "초록", "진한 초록",
            "보라", "밝은 보라", "자홍", "분홍", "청록", "밝은 청록", "갈색", "황갈색",
            "아주 연한 파랑", "연한 파랑", "아주 연한 초록", "연한 초록", "연한 노랑", "살구색", "연한 주황", "연한 보라"
        };
        internal static float ParseSize(string text)
        {
            float size;
            string value = (text ?? "").Trim();
            if ((!float.TryParse(value, NumberStyles.Float, CultureInfo.CurrentCulture, out size) &&
                 !float.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out size)) ||
                float.IsNaN(size) || float.IsInfinity(size) || size < 1 || size > 400)
                throw new ArgumentException("글자 크기는 1~400 pt 사이의 숫자로 입력해 주세요.");
            return size;
        }
        internal static string Hex(Color color) { return color.R.ToString("X2") + color.G.ToString("X2") + color.B.ToString("X2"); }
        internal static Color ParseColor(string text)
        {
            string value = (text ?? "").Trim().TrimStart('#');
            int rgb;
            if (value.Length != 6 || !int.TryParse(value, NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out rgb))
                throw new ArgumentException("색상 코드는 000000처럼 여섯 자리로 입력해 주세요.");
            return Color.FromArgb(255, (rgb >> 16) & 255, (rgb >> 8) & 255, rgb & 255);
        }
    }

    public partial class Connect
    {
        private object ribbonUI;
        private readonly Func<NumberLabelSettings> readNumberSettings;
        private readonly Action<NumberLabelSettings> writeNumberSettings;
        public Connect() : this(NumberLabelPreferences.Load, NumberLabelPreferences.Save) { }
        internal Connect(Func<NumberLabelSettings> read, Action<NumberLabelSettings> write)
        {
            readNumberSettings = read; writeNumberSettings = write;
        }
        public void OnRibbonLoad(object ribbon) { ribbonUI = ribbon; }
        public int GetNumberFontCount(object control) { return NumberRibbonSettings.Fonts.Length; }
        public string GetNumberFontLabel(object control, int index) { return NumberRibbonSettings.Fonts[index]; }
        public int GetNumberFontIndex(object control)
        {
            string selectedFont = readNumberSettings().FontName;
            int index = Array.FindIndex(NumberRibbonSettings.Fonts, f => string.Equals(f, selectedFont, StringComparison.OrdinalIgnoreCase));
            return index >= 0 ? index : Math.Max(0, Array.IndexOf(NumberRibbonSettings.Fonts, "Arial"));
        }
        public void SetNumberFont(object control, string selectedId, int index)
        {
            ChangeNumberSettings(s => s.FontName = NumberRibbonSettings.Fonts[index]);
        }
        public string GetNumberFontSize(object control) { return readNumberSettings().FontSize.ToString("0.###", CultureInfo.CurrentCulture); }
        public void SetNumberFontSize(object control, string text) { ChangeNumberSettings(s => s.FontSize = NumberRibbonSettings.ParseSize(text)); }
        public int GetNumberColorCount(object control) { return NumberRibbonSettings.Colors.Length; }
        public string GetNumberColorLabel(object control, int index) { return NumberRibbonSettings.ColorNames[index] + " (#" + NumberRibbonSettings.Hex(NumberRibbonSettings.Colors[index]) + ")"; }
        public object GetNumberColorItemImage(object control, int index) { return RibbonIcons.ColorSwatch(NumberRibbonSettings.Colors[index]); }
        public object GetNumberColorImage(object control) { return RibbonIcons.ColorSwatch(readNumberSettings().TextColor); }
        public void SetNumberColor(object control, string selectedId, int index) { ChangeNumberSettings(s => s.TextColorArgb = NumberRibbonSettings.Colors[index].ToArgb()); }
        public string GetNumberColorHex(object control) { return NumberRibbonSettings.Hex(readNumberSettings().TextColor); }
        public void SetNumberColorHex(object control, string text) { ChangeNumberSettings(s => s.TextColorArgb = NumberRibbonSettings.ParseColor(text).ToArgb()); }
        private void ChangeNumberSettings(Action<NumberLabelSettings> change)
        {
            try
            {
                Run(delegate { NumberLabelSettings settings = readNumberSettings().Copy(); change(settings); writeNumberSettings(settings); });
            }
            finally
            {
                if (ribbonUI != null)
                    foreach (string id in new[] { "labNumberFont", "labNumberFontSize", "labNumberColor", "labNumberColorHex" })
                        ((dynamic)ribbonUI).InvalidateControl(id);
            }
        }
    }
}


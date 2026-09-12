using System;
using System.Collections.Generic;

namespace LabPhotoTools
{
    public static partial class Layout
    {
        // The array occupies at most 80% of each slide dimension: 10% per edge.
        // Equal-height rows preserve every photograph's aspect ratio. Fill a
        // fixed number of columns before starting the next row (5/5, 4/4/3,
        // etc.).
        // Only the final row can be incomplete; every row shares the left edge.
        public static List<PhotoBox> SmartArrange(IList<PhotoBox> photos, double slideWidth, double slideHeight)
        {
            List<PhotoBox> input = ValidateAndClone(photos, 0, 0, slideWidth, slideHeight);
            List<PhotoBox> ordered = new List<PhotoBox>();
            foreach (Row row in DetectRows(input)) ordered.AddRange(row.Photos);
            const double relativeGap = 0.06;
            int bestColumns = PreferredColumns(ordered.Count);
            if (bestColumns == 0)
            {
                double bestScore = double.PositiveInfinity;
                for (int columns = 1; columns <= ordered.Count; columns++)
                {
                    int rows = (ordered.Count + columns - 1) / columns;
                    double maxWidth = 0;
                    int index = 0;
                    for (int r = 0; r < rows; r++)
                    {
                        int count = Math.Min(columns, ordered.Count - index);
                        double rowWidth = (count - 1) * relativeGap;
                        for (int c = 0; c < count; c++, index++) rowWidth += ordered[index].Width / ordered[index].Height;
                        maxWidth = Math.Max(maxWidth, rowWidth);
                    }
                    double height = rows + (rows - 1) * relativeGap;
                    double score = Math.Abs(Math.Log((maxWidth / height) / (slideWidth / slideHeight)));
                    if (score < bestScore) { bestScore = score; bestColumns = columns; }
                }
            }
            int bestRows = (ordered.Count + bestColumns - 1) / bestColumns;
            int cursor = 0;
            double widest = 0;
            for (int r = 0; r < bestRows; r++)
            {
                int count = Math.Min(bestColumns, ordered.Count - cursor);
                double rowWidth = (count - 1) * relativeGap;
                for (int c = 0; c < count; c++, cursor++) rowWidth += ordered[cursor].Width / ordered[cursor].Height;
                widest = Math.Max(widest, rowWidth);
            }
            double totalHeight = bestRows + (bestRows - 1) * relativeGap;
            double scale = Math.Min(slideWidth * 0.8 / widest, slideHeight * 0.8 / totalHeight);
            double top = (slideHeight - totalHeight * scale) / 2;
            double arrayLeft = (slideWidth - widest * scale) / 2;
            cursor = 0;
            for (int r = 0; r < bestRows; r++)
            {
                int count = Math.Min(bestColumns, ordered.Count - cursor);
                double left = arrayLeft;
                for (int c = 0; c < count; c++, cursor++)
                {
                    PhotoBox photo = ordered[cursor];
                    double ratio = photo.Width / photo.Height;
                    photo.Left = left;
                    photo.Top = top;
                    photo.Width = ratio * scale;
                    photo.Height = scale;
                    left += photo.Width + relativeGap * scale;
                }
                top += (1 + relativeGap) * scale;
            }
            ValidateBounds(ordered, slideWidth, slideHeight);
            return ordered;
        }

        // These layouts are deliberately independent of slide orientation so
        // small, common photo counts have a predictable compact arrangement.
        // Larger counts continue to choose the grid closest to the slide and
        // the photographs' combined aspect ratio.
        private static int PreferredColumns(int count)
        {
            switch (count)
            {
                case 4: return 2;
                case 5: case 6: return 3;
                case 7: case 8: return 4;
                case 9: return 3;
                case 10: return 5;
                case 11: case 12: case 13: case 14: return 4;
                case 15: return 5;
                case 16: return 4;
                default: return 0;
            }
        }
    }
}


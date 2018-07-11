using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Media;

namespace YmlToExcelTool
{
    public class TextHelper
    {
        static readonly TextBlock TextBlock = new TextBlock() { Width = 1920 };
        public static bool CompareTextLength(string text1, string text2)
        {
            Typeface typeface = new Typeface(
                TextBlock.FontFamily,
                TextBlock.FontStyle,
                TextBlock.FontWeight,
                TextBlock.FontStretch);

            FormattedText formattedText1 = new FormattedText(
                text1,
                System.Threading.Thread.CurrentThread.CurrentCulture,
                TextBlock.FlowDirection,
                typeface,
                TextBlock.FontSize,
                TextBlock.Foreground);
            FormattedText formattedText2 = new FormattedText(
                text2,
                System.Threading.Thread.CurrentThread.CurrentCulture,
                TextBlock.FlowDirection,
                typeface,
                TextBlock.FontSize,
                TextBlock.Foreground);

            return formattedText1.Width < formattedText2.Width;
        }
    }
}

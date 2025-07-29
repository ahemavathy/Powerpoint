using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace PowerPointGenerator.Utilities
{
    /// <summary>
    /// Helper class for creating presentation themes
    /// </summary>
    public static class ThemeHelper
    {
        /// <summary>
        /// Creates a default theme for the presentation
        /// </summary>
        /// <returns>Theme object</returns>
        public static A.Theme CreateDefaultTheme()
        {
            var theme = new A.Theme() { Name = "Office Theme" };

            var themeElements = new A.ThemeElements();

            // Color scheme
            var colorScheme = new A.ColorScheme() { Name = "Office" };
            colorScheme.Append(new A.Dark1Color(new A.SystemColor() { Val = A.SystemColorValues.WindowText, LastColor = "000000" }));
            colorScheme.Append(new A.Light1Color(new A.SystemColor() { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" }));
            colorScheme.Append(new A.Dark2Color(new A.RgbColorModelHex() { Val = "44546A" }));
            colorScheme.Append(new A.Light2Color(new A.RgbColorModelHex() { Val = "E7E6E6" }));
            colorScheme.Append(new A.Accent1Color(new A.RgbColorModelHex() { Val = "4472C4" }));
            colorScheme.Append(new A.Accent2Color(new A.RgbColorModelHex() { Val = "E15759" }));
            colorScheme.Append(new A.Accent3Color(new A.RgbColorModelHex() { Val = "70AD47" }));
            colorScheme.Append(new A.Accent4Color(new A.RgbColorModelHex() { Val = "FFC000" }));
            colorScheme.Append(new A.Accent5Color(new A.RgbColorModelHex() { Val = "5B9BD5" }));
            colorScheme.Append(new A.Accent6Color(new A.RgbColorModelHex() { Val = "FF6600" }));
            colorScheme.Append(new A.Hyperlink(new A.RgbColorModelHex() { Val = "0563C1" }));
            colorScheme.Append(new A.FollowedHyperlinkColor(new A.RgbColorModelHex() { Val = "954F72" }));

            // Font scheme
            var fontScheme = new A.FontScheme() { Name = "Office" };

            var majorFont = new A.MajorFont();
            majorFont.Append(new A.LatinFont() { Typeface = "Calibri Light", Panose = "020F0302020204030204" });
            majorFont.Append(new A.EastAsianFont() { Typeface = "" });
            majorFont.Append(new A.ComplexScriptFont() { Typeface = "" });

            var minorFont = new A.MinorFont();
            minorFont.Append(new A.LatinFont() { Typeface = "Calibri", Panose = "020F0502020204030204" });
            minorFont.Append(new A.EastAsianFont() { Typeface = "" });
            minorFont.Append(new A.ComplexScriptFont() { Typeface = "" });

            fontScheme.Append(majorFont);
            fontScheme.Append(minorFont);

            // Format scheme
            var formatScheme = new A.FormatScheme() { Name = "Office" };

            var fillStyleList = new A.FillStyleList();
            fillStyleList.Append(new A.SolidFill(new A.SchemeColor() { Val = A.SchemeColorValues.PhColor }));
            fillStyleList.Append(new A.GradientFill(
                new A.GradientStopList(
                    new A.GradientStop() { Position = 0 },
                    new A.GradientStop() { Position = 100000 }),
                new A.LinearGradientFill() { Angle = 5400000, Scaled = true }));
            fillStyleList.Append(new A.GradientFill(
                new A.GradientStopList(
                    new A.GradientStop() { Position = 0 },
                    new A.GradientStop() { Position = 100000 }),
                new A.LinearGradientFill() { Angle = 5400000, Scaled = true }));

            var lineStyleList = new A.LineStyleList();
            lineStyleList.Append(new A.Outline() { Width = 9525 });
            lineStyleList.Append(new A.Outline() { Width = 25400 });
            lineStyleList.Append(new A.Outline() { Width = 38100 });

            var effectStyleList = new A.EffectStyleList();
            effectStyleList.Append(new A.EffectStyle(new A.EffectList()));
            effectStyleList.Append(new A.EffectStyle(new A.EffectList()));
            effectStyleList.Append(new A.EffectStyle(new A.EffectList()));

            var backgroundFillStyleList = new A.BackgroundFillStyleList();
            backgroundFillStyleList.Append(new A.SolidFill(new A.SchemeColor() { Val = A.SchemeColorValues.PhColor }));
            backgroundFillStyleList.Append(new A.GradientFill());
            backgroundFillStyleList.Append(new A.GradientFill());

            formatScheme.Append(fillStyleList);
            formatScheme.Append(lineStyleList);
            formatScheme.Append(effectStyleList);
            formatScheme.Append(backgroundFillStyleList);

            themeElements.Append(colorScheme);
            themeElements.Append(fontScheme);
            themeElements.Append(formatScheme);

            theme.Append(themeElements);
            theme.Append(new A.ObjectDefaults());
            theme.Append(new A.ExtraColorSchemeList());

            return theme;
        }
    }
}

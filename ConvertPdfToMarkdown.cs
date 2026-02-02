using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Core;

namespace AllToMarkdown
{
    internal class ConvertPdf
    {
        /// <summary>
        /// Extracts text from PDF data using PdfPig and formats it as Markdown with page separators.
        /// </summary>
        public static string ConvertPdfToMarkdown(Stream data)
        {
            var markdownBuilder = new StringBuilder();
            var pageSeparator = "\n\n---\n\n"; // Page separator
            var previousPageEnd = new TextMetrics();

            try
            {
                using (var pdfDocument = PdfDocument.Open(data))
                {
                    int totalPages = pdfDocument.NumberOfPages;
                    var pages = pdfDocument.GetPages().ToList();

                    for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++)
                    {
                        var page = pages[pageIndex];
                        var pageNumber = pageIndex + 1;

                        // Progress log
                        Debug.WriteLine($"Processing page {pageNumber}/{totalPages}");

                        // Extract and analyze page content
                        var pageContent = ExtractStructuredPageContent(page, previousPageEnd);

                        // Convert to Markdown
                        var pageMarkdown = FormatPageAsMarkdown(pageContent, pageNumber);
                        markdownBuilder.Append(pageMarkdown);

                        // Save metrics for cross-page context
                        previousPageEnd = pageContent.Metrics;

                        // Add separator between pages (except last)
                        if (pageIndex < pages.Count - 1)
                        {
                            markdownBuilder.Append(pageSeparator);
                        }
                    }
                }

                // Global document post-processing
                var finalMarkdown = PostProcessMarkdown(markdownBuilder.ToString());
                return finalMarkdown;
            }
            catch (Exception ex)
            {
                throw new PdfConversionException($"Error converting PDF to Markdown: {ex.Message}", ex);
            }
        }

        #region Support Structures

        private class PageSize
        {
            public double Width { get; set; }
            public double Height { get; set; }
        }

        private class PageContent
        {
            public List<TextElement> Elements { get; set; } = new List<TextElement>();
            public TextMetrics Metrics { get; set; } = new TextMetrics();
            public List<DetectedTable> Tables { get; set; } = new List<DetectedTable>();
            public PageSize Size { get; set; }
        }

        private class TextElement
        {
            public string Text { get; set; }
            public TextFormat Format { get; set; }
            public TextPosition Position { get; set; }
            public ElementType Type { get; set; }
            public double Confidence { get; set; } = 1.0;
            public int IndentLevel { get; set; }
            public List<TextElement> Children { get; set; } = new List<TextElement>();
            public bool IsContinuation { get; set; }
        }

        private class TextFormat
        {
            public string FontName { get; set; }
            public double FontSize { get; set; }
            public bool IsBold { get; set; }
            public bool IsItalic { get; set; }
            public bool IsUnderlined { get; set; }
            public RGBColor Color { get; set; }
        }

        private class TextPosition
        {
            public double X { get; set; }
            public double Y { get; set; }
            public double Width { get; set; }
            public double Height { get; set; }
        }

        private class TextMetrics
        {
            public double MostCommonFontSize { get; set; }
            public double HeadingFontSize { get; set; }
            public double SubheadingFontSize { get; set; }
            public double NormalFontSize { get; set; }
            public Dictionary<string, int> FontUsage { get; set; } = new Dictionary<string, int>();
            public double AverageLineHeight { get; set; }
            public double LeftMargin { get; set; }
        }

        private class DetectedTable
        {
            public List<TableRow> Rows { get; set; } = new List<TableRow>();
            public TextPosition Bounds { get; set; }
            public int ColumnCount { get; set; }
        }

        private class TableRow
        {
            public List<TableCell> Cells { get; set; } = new List<TableCell>();
            public bool IsHeader { get; set; }
        }

        private class TableCell
        {
            public string Text { get; set; }
            public int ColSpan { get; set; } = 1;
            public int RowSpan { get; set; } = 1;
        }

        private class RGBColor
        {
            public byte R { get; set; }
            public byte G { get; set; }
            public byte B { get; set; }

            public RGBColor() { }

            public RGBColor(byte r, byte g, byte b)
            {
                R = r;
                G = g;
                B = b;
            }

            public override bool Equals(object obj)
            {
                return obj is RGBColor color &&
                       R == color.R &&
                       G == color.G &&
                       B == color.B;
            }

            public override int GetHashCode()
            {
                return HashCode.Combine(R, G, B);
            }
        }

        private enum ElementType
        {
            Unknown,
            Heading1,
            Heading2,
            Heading3,
            Heading4,
            Paragraph,
            ListItem,
            NumberedListItem,
            Table,
            CodeBlock,
            BlockQuote,
            HorizontalRule,
            ImageCaption,
            Footer,
            Header,
            Caption,
            PageNumber,
            Footnote
        }

        #endregion

        #region Support Classes for Internal Processing

        private class TextWord
        {
            public List<Letter> Letters { get; set; } = new List<Letter>();

            public bool IsBold
            {
                get
                {
                    if (!Letters.Any()) return false;
                    var fontName = Letters.First().FontName ?? "";
                    return fontName.IndexOf("bold", StringComparison.OrdinalIgnoreCase) >= 0 ||
                           fontName.IndexOf("black", StringComparison.OrdinalIgnoreCase) >= 0 ||
                           fontName.IndexOf("heavy", StringComparison.OrdinalIgnoreCase) >= 0 ||
                           fontName.IndexOf("700", StringComparison.OrdinalIgnoreCase) >= 0 ||
                           fontName.IndexOf("800", StringComparison.OrdinalIgnoreCase) >= 0 ||
                           fontName.IndexOf("900", StringComparison.OrdinalIgnoreCase) >= 0;
                }
            }

            public bool IsItalic
            {
                get
                {
                    if (!Letters.Any()) return false;
                    var fontName = Letters.First().FontName ?? "";
                    return fontName.IndexOf("italic", StringComparison.OrdinalIgnoreCase) >= 0 ||
                           fontName.IndexOf("oblique", StringComparison.OrdinalIgnoreCase) >= 0 ||
                           fontName.IndexOf("italic", StringComparison.OrdinalIgnoreCase) >= 0;
                }
            }

            public double AvgFontSize => Letters.Any() ? Letters.Average(l => l.FontSize) : 0;

            public PdfRectangle BoundingBox => CalculateBoundingBox();

            public RGBColor Color
            {
                get
                {
                    if (!Letters.Any()) return new RGBColor(0, 0, 0);

                    var firstLetter = Letters.First();
                    var color = firstLetter.Color;

                    if (color != null)
                    {
                        // Try to get RGB values using dynamic or manual extraction
                        try
                        {
                            // Method 1: Try to access properties via dynamic
                            dynamic dynamicColor = color;
                            byte r = 0, g = 0, b = 0;

                            // Try different property names that might exist
                            if (HasProperty(dynamicColor, "R")) r = dynamicColor.R;
                            else if (HasProperty(dynamicColor, "Red")) r = dynamicColor.Red;

                            if (HasProperty(dynamicColor, "G")) g = dynamicColor.G;
                            else if (HasProperty(dynamicColor, "Green")) g = dynamicColor.Green;

                            if (HasProperty(dynamicColor, "B")) b = dynamicColor.B;
                            else if (HasProperty(dynamicColor, "Blue")) b = dynamicColor.Blue;

                            return new RGBColor(r, g, b);
                        }
                        catch
                        {
                            // If dynamic fails, try reflection
                            var colorType = color.GetType();
                            var properties = colorType.GetProperties();

                            byte r = 0, g = 0, b = 0;

                            foreach (var prop in properties)
                            {
                                var propName = prop.Name.ToUpper();
                                if (propName == "R" || propName == "RED")
                                    r = Convert.ToByte(prop.GetValue(color) ?? 0);
                                else if (propName == "G" || propName == "GREEN")
                                    g = Convert.ToByte(prop.GetValue(color) ?? 0);
                                else if (propName == "B" || propName == "BLUE")
                                    b = Convert.ToByte(prop.GetValue(color) ?? 0);
                            }

                            return new RGBColor(r, g, b);
                        }
                    }

                    return new RGBColor(0, 0, 0); // Default black
                }
            }

            private static bool HasProperty(dynamic obj, string propertyName)
            {
                try
                {
                    var value = obj[propertyName];
                    return true;
                }
                catch
                {
                    return false;
                }
            }

            private PdfRectangle CalculateBoundingBox()
            {
                if (!Letters.Any()) return new PdfRectangle(0, 0, 0, 0);

                // Use GlyphRectangle for accurate bounding box
                var firstGlyph = Letters.First().GlyphRectangle;
                double left = firstGlyph.Left;
                double bottom = firstGlyph.Bottom;
                double right = firstGlyph.Right;
                double top = firstGlyph.Top;

                foreach (var letter in Letters.Skip(1))
                {
                    var glyph = letter.GlyphRectangle;
                    left = Math.Min(left, glyph.Left);
                    bottom = Math.Min(bottom, glyph.Bottom);
                    right = Math.Max(right, glyph.Right);
                    top = Math.Max(top, glyph.Top);
                }

                return new PdfRectangle(left, bottom, right, top);
            }

            public override string ToString() =>
                string.Join("", Letters.Select(l => l.Value));
        }

        private class TextLine
        {
            public List<TextWord> Words { get; set; } = new List<TextWord>();
            public PdfRectangle BoundingBox => CalculateBoundingBox();
            public double AvgFontSize => Words.Any() ? Words.Average(w => w.AvgFontSize) : 0;
            public string MostCommonFont => GetMostCommonFont();
            public RGBColor MostCommonColor => GetMostCommonColor();
            public double LineHeight => CalculateLineHeight();

            private PdfRectangle CalculateBoundingBox()
            {
                if (!Words.Any()) return new PdfRectangle(0, 0, 0, 0);

                var firstBox = Words.First().BoundingBox;
                double left = firstBox.Left;
                double bottom = firstBox.Bottom;
                double right = firstBox.Right;
                double top = firstBox.Top;

                foreach (var word in Words.Skip(1))
                {
                    var box = word.BoundingBox;
                    left = Math.Min(left, box.Left);
                    bottom = Math.Min(bottom, box.Bottom);
                    right = Math.Max(right, box.Right);
                    top = Math.Max(top, box.Top);
                }

                return new PdfRectangle(left, bottom, right, top);
            }

            private string GetMostCommonFont()
            {
                var fonts = Words
                    .SelectMany(w => w.Letters)
                    .Where(l => !string.IsNullOrEmpty(l.FontName))
                    .Select(l => l.FontName);

                if (!fonts.Any()) return "Unknown";

                return fonts
                    .GroupBy(f => f)
                    .OrderByDescending(g => g.Count())
                    .First().Key;
            }

            private RGBColor GetMostCommonColor()
            {
                var colors = Words.Select(w => w.Color);
                if (!colors.Any()) return new RGBColor(0, 0, 0);

                var colorGroups = colors
                    .GroupBy(c => c)
                    .OrderByDescending(g => g.Count())
                    .ToList();

                return colorGroups.First().Key;
            }

            private double CalculateLineHeight()
            {
                if (!Words.Any()) return 0;
                var maxTop = Words.Max(w => w.BoundingBox.Top);
                var minBottom = Words.Min(w => w.BoundingBox.Bottom);
                return maxTop - minBottom;
            }

            public override string ToString() =>
                string.Join(" ", Words.Select(w => w.ToString()));
        }

        private class PdfConversionException : Exception
        {
            public PdfConversionException(string message, Exception inner)
                : base(message, inner) { }
        }

        #endregion

        #region Structured Extraction

        private static PageContent ExtractStructuredPageContent(UglyToad.PdfPig.Content.Page page, TextMetrics previousPageMetrics)
        {
            var content = new PageContent();

            // Extract all letters
            var letters = page.Letters.ToList();

            if (!letters.Any())
                return content;

            // Calculate page metrics
            content.Metrics = CalculatePageMetrics(letters, previousPageMetrics, page.Width);
            content.Size = new PageSize
            {
                Width = page.Width,
                Height = page.Height
            };

            // Group letters into words
            var words = GroupLettersIntoWords(letters);

            // Group words into lines
            var lines = GroupWordsIntoLines(words, content.Metrics);

            // Detect tables before normal processing
            content.Tables = DetectTables(lines, content.Size.Width);

            // Filter lines that are part of tables
            var nonTableLines = FilterOutTableLines(lines, content.Tables);

            // Analyze hierarchical structure
            var elements = AnalyzeDocumentStructure(nonTableLines, content.Metrics);

            // Identify element types
            elements = ClassifyElements(elements, content.Metrics);

            // Detect list nesting
            elements = DetectListNesting(elements);

            // Group related elements
            elements = GroupRelatedElements(elements);

            content.Elements = elements;

            return content;
        }

        private static TextMetrics CalculatePageMetrics(List<Letter> letters, TextMetrics previousMetrics, double pageWidth)
        {
            var metrics = new TextMetrics();

            if (!letters.Any())
                return previousMetrics;

            // Calculate font size distribution
            var fontSizeGroups = letters
                .Where(l => l.FontSize > 0)
                .GroupBy(l => Math.Round(l.FontSize, 1))
                .OrderByDescending(g => g.Count())
                .ToList();

            if (fontSizeGroups.Any())
            {
                metrics.MostCommonFontSize = fontSizeGroups[0].Key;
                metrics.NormalFontSize = fontSizeGroups[0].Key;

                // Identify heading sizes
                if (fontSizeGroups.Count > 1)
                {
                    var largeFonts = fontSizeGroups.Where(g => g.Key > metrics.NormalFontSize * 1.3);
                    if (largeFonts.Any())
                    {
                        metrics.HeadingFontSize = largeFonts.Max(g => g.Key);

                        var mediumFonts = fontSizeGroups.Where(g =>
                            g.Key > metrics.NormalFontSize * 1.1 &&
                            g.Key < metrics.HeadingFontSize);
                        if (mediumFonts.Any())
                        {
                            metrics.SubheadingFontSize = mediumFonts.Max(g => g.Key);
                        }
                    }
                }
            }

            // Font usage statistics
            foreach (var letter in letters)
            {
                var fontKey = $"{letter.FontName}_{Math.Round(letter.FontSize, 1)}";
                if (!metrics.FontUsage.ContainsKey(fontKey))
                    metrics.FontUsage[fontKey] = 0;
                metrics.FontUsage[fontKey]++;
            }

            // Calculate average line height based on letter positions
            var yPositions = letters.Select(l => l.GlyphRectangle.Top).Distinct().OrderByDescending(y => y).ToList();
            if (yPositions.Count > 1)
            {
                var gaps = new List<double>();
                for (int i = 1; i < yPositions.Count; i++)
                {
                    var gap = yPositions[i - 1] - yPositions[i]; // Top to bottom
                    if (gap > 0 && gap < metrics.NormalFontSize * 3) // Reasonable line gap
                        gaps.Add(gap);
                }
                metrics.AverageLineHeight = gaps.Any() ? gaps.Average() : metrics.NormalFontSize * 1.2;
            }
            else
            {
                metrics.AverageLineHeight = metrics.NormalFontSize * 1.2;
            }

            // Estimate left margin from the leftmost position of letters
            var leftPositions = letters.Select(l => l.GlyphRectangle.Left)
                                       .Where(x => x > 0 && x < pageWidth * 0.8) // Exclude far right
                                       .ToList();
            metrics.LeftMargin = leftPositions.Any() ? leftPositions.Min() : 50;

            // Use previous metrics if these are not valid
            if (metrics.HeadingFontSize == 0 && previousMetrics.HeadingFontSize > 0)
                metrics.HeadingFontSize = previousMetrics.HeadingFontSize;

            if (metrics.SubheadingFontSize == 0 && previousMetrics.SubheadingFontSize > 0)
                metrics.SubheadingFontSize = previousMetrics.SubheadingFontSize;

            return metrics;
        }

        private static List<TextWord> GroupLettersIntoWords(List<Letter> letters)
        {
            var words = new List<TextWord>();

            if (!letters.Any())
                return words;

            // Sort letters by position (top to bottom, left to right)
            var orderedLetters = letters
                .OrderByDescending(l => l.GlyphRectangle.Top) // Sort by top position
                .ThenBy(l => l.GlyphRectangle.Left)
                .ToList();

            var currentWord = new TextWord();
            Letter previousLetter = null;

            foreach (var letter in orderedLetters)
            {
                bool shouldStartNewWord = false;

                if (previousLetter != null && currentWord.Letters.Any())
                {
                    // Check if letters belong to same word
                    var currentTop = letter.GlyphRectangle.Top;
                    var previousTop = previousLetter.GlyphRectangle.Top;
                    var verticalDiff = Math.Abs(currentTop - previousTop);

                    var currentLeft = letter.GlyphRectangle.Left;
                    var previousRight = previousLetter.GlyphRectangle.Right;
                    var horizontalGap = currentLeft - previousRight;

                    // Letters belong to same word if:
                    // 1. They're on roughly the same line (vertical diff less than 40% of font size)
                    // 2. Horizontal gap is less than space character width (approx 0.4 * font size)
                    shouldStartNewWord = verticalDiff > letter.FontSize * 0.4 ||
                                        horizontalGap > letter.FontSize * 0.6;
                }

                if (shouldStartNewWord && currentWord.Letters.Any())
                {
                    words.Add(currentWord);
                    currentWord = new TextWord();
                }

                currentWord.Letters.Add(letter);
                previousLetter = letter;
            }

            if (currentWord.Letters.Any())
                words.Add(currentWord);

            return words;
        }

        private static List<TextLine> GroupWordsIntoLines(List<TextWord> words, TextMetrics metrics)
        {
            var lines = new List<TextLine>();

            if (!words.Any())
                return lines;

            // Group words by their vertical position (Y coordinate)
            var yTolerance = metrics.AverageLineHeight * 0.3;

            // Create a dictionary to group words by their "line band"
            var lineGroups = new Dictionary<int, List<TextWord>>();

            foreach (var word in words)
            {
                var wordTop = (int)Math.Round(word.BoundingBox.Top / yTolerance);

                if (!lineGroups.ContainsKey(wordTop))
                    lineGroups[wordTop] = new List<TextWord>();

                lineGroups[wordTop].Add(word);
            }

            // Create lines from groups, sorted top to bottom
            foreach (var group in lineGroups.OrderByDescending(g => g.Key))
            {
                var line = new TextLine();
                // Sort words in line from left to right
                var sortedWords = group.Value.OrderBy(w => w.BoundingBox.Left).ToList();
                line.Words.AddRange(sortedWords);
                lines.Add(line);
            }

            return lines;
        }

        private static List<TextElement> AnalyzeDocumentStructure(List<TextLine> lines, TextMetrics metrics)
        {
            var elements = new List<TextElement>();

            for (int i = 0; i < lines.Count; i++)
            {
                var line = lines[i];
                var text = line.ToString().Trim();

                if (string.IsNullOrWhiteSpace(text))
                    continue;

                // Check if this is a continuation of previous line
                bool isContinuation = i > 0 &&
                    lines[i - 1].BoundingBox.Left > metrics.LeftMargin * 0.8 &&
                    line.BoundingBox.Left > metrics.LeftMargin * 0.8 &&
                    Math.Abs(lines[i - 1].BoundingBox.Top - line.BoundingBox.Top) < metrics.AverageLineHeight * 1.5 &&
                    !lines[i - 1].ToString().TrimEnd().EndsWith(".") &&
                    !lines[i - 1].ToString().TrimEnd().EndsWith("!") &&
                    !lines[i - 1].ToString().TrimEnd().EndsWith("?") &&
                    !lines[i - 1].ToString().TrimEnd().EndsWith(":") &&
                    text.Length > 0 &&
                    !char.IsUpper(text[0]) &&
                    !IsListItem(text);

                var element = new TextElement
                {
                    Text = text,
                    Position = new TextPosition
                    {
                        X = line.BoundingBox.Left,
                        Y = line.BoundingBox.Top,
                        Width = line.BoundingBox.Width,
                        Height = line.BoundingBox.Height
                    },
                    Format = new TextFormat
                    {
                        FontName = line.MostCommonFont,
                        FontSize = line.AvgFontSize,
                        IsBold = line.Words.Any(w => w.IsBold),
                        IsItalic = line.Words.Any(w => w.IsItalic),
                        IsUnderlined = DetectUnderline(line),
                        Color = line.MostCommonColor
                    },
                    IsContinuation = isContinuation
                };

                elements.Add(element);
            }

            return elements;
        }

        private static List<TextElement> ClassifyElements(List<TextElement> elements, TextMetrics metrics)
        {
            var classified = new List<TextElement>();

            for (int i = 0; i < elements.Count; i++)
            {
                var element = elements[i];
                var text = element.Text.Trim();

                if (string.IsNullOrWhiteSpace(text) || element.IsContinuation)
                    continue;

                var fontSize = element.Format.FontSize;
                var isBold = element.Format.IsBold;
                var isItalic = element.Format.IsItalic;
                var xPos = element.Position.X;

                // 1. Detect page numbers (often centered, small font, at bottom)
                if (fontSize < metrics.NormalFontSize * 0.9 &&
                    xPos > metrics.LeftMargin * 3 &&
                    Regex.IsMatch(text, @"^\d+$"))
                {
                    element.Type = ElementType.PageNumber;
                    element.Confidence = 0.9;
                }
                // 2. Detect headings
                else if (fontSize >= metrics.HeadingFontSize * 0.85 && fontSize <= metrics.HeadingFontSize * 1.15)
                {
                    element.Type = isBold ? ElementType.Heading1 : ElementType.Heading2;
                    element.Confidence = 0.9;
                }
                else if (fontSize >= metrics.SubheadingFontSize * 0.85 && fontSize <= metrics.SubheadingFontSize * 1.15)
                {
                    element.Type = isBold ? ElementType.Heading2 : ElementType.Heading3;
                    element.Confidence = 0.8;
                }
                else if (fontSize > metrics.NormalFontSize * 1.2 && isBold)
                {
                    element.Type = ElementType.Heading3;
                    element.Confidence = 0.7;
                }
                else if (fontSize > metrics.NormalFontSize && isBold)
                {
                    element.Type = ElementType.Heading4;
                    element.Confidence = 0.6;
                }
                // 3. Detect lists
                else if (IsListItem(text))
                {
                    var listType = GetListItemType(text);
                    element.Type = listType;
                    element.IndentLevel = CalculateIndentLevel(xPos, metrics.LeftMargin);
                    element.Confidence = 0.85;
                    element.Text = CleanListItemText(text);
                }
                // 4. Detect code blocks (monospace font, often indented)
                else if (IsMonospaceFont(element.Format.FontName))
                {
                    element.Type = ElementType.CodeBlock;
                    element.Confidence = 0.75;
                }
                // 5. Detect blockquotes (indented from left)
                else if (xPos > metrics.LeftMargin * 1.5 && xPos < metrics.LeftMargin * 4)
                {
                    element.Type = ElementType.BlockQuote;
                    element.Confidence = 0.7;
                }
                // 6. Detect horizontal rules
                else if (IsHorizontalRule(text))
                {
                    element.Type = ElementType.HorizontalRule;
                    element.Confidence = 1.0;
                }
                // 7. Detect footnotes (small font, often starts with numbers or symbols)
                else if (fontSize < metrics.NormalFontSize * 0.9 &&
                        (text.StartsWith("[") || Regex.IsMatch(text, @"^[†‡*#]") ||
                         Regex.IsMatch(text, @"^\d+[\.\)]")))
                {
                    element.Type = ElementType.Footnote;
                    element.Confidence = 0.8;
                }
                // 8. Normal paragraph
                else
                {
                    element.Type = ElementType.Paragraph;
                    element.Confidence = 1.0;
                }

                classified.Add(element);
            }

            return classified;
        }

        #endregion

        #region Advanced Table Detection

        private static List<DetectedTable> DetectTables(List<TextLine> lines, double pageWidth)
        {
            var tables = new List<DetectedTable>();

            if (lines.Count < 2)
                return tables;

            // Group rows that might form a table
            var tableCandidates = new List<List<TextLine>>();
            var currentCandidate = new List<TextLine>();

            for (int i = 0; i < lines.Count; i++)
            {
                var line = lines[i];
                var text = line.ToString().Trim();

                if (string.IsNullOrWhiteSpace(text))
                    continue;

                if (IsLikelyTableRow(line, pageWidth))
                {
                    currentCandidate.Add(line);

                    // If next row is not a table or we're at the end, finalize candidate
                    if (i == lines.Count - 1 ||
                        !IsLikelyTableRow(lines[i + 1], pageWidth) ||
                        IsSignificantVerticalGap(line, lines[i + 1]))
                    {
                        if (currentCandidate.Count >= 2) // At least 2 rows for a table
                        {
                            tableCandidates.Add(new List<TextLine>(currentCandidate));
                        }
                        currentCandidate.Clear();
                    }
                }
                else if (currentCandidate.Count > 0)
                {
                    // Non-table line breaks the table
                    if (currentCandidate.Count >= 2)
                    {
                        tableCandidates.Add(new List<TextLine>(currentCandidate));
                    }
                    currentCandidate.Clear();
                }
            }

            // Analyze each candidate
            foreach (var candidate in tableCandidates)
            {
                var table = AnalyzeTableStructure(candidate, pageWidth);
                if (table.Rows.Count >= 2)
                {
                    tables.Add(table);
                }
            }

            return tables;
        }

        private static bool IsLikelyTableRow(TextLine line, double pageWidth)
        {
            var text = line.ToString().Trim();

            if (string.IsNullOrWhiteSpace(text) || text.Length > 200)
                return false;

            var words = line.Words;

            // Criterion 1: Contains pipe characters (common in text tables)
            if (text.Contains("|") && text.Count(c => c == '|') >= 2)
                return true;

            // Criterion 2: Contains tabular alignment (multiple words with regular spacing)
            if (words.Count >= 3)
            {
                // Calculate distances between words
                var gaps = new List<double>();
                for (int i = 1; i < words.Count; i++)
                {
                    var gap = words[i].BoundingBox.Left - words[i - 1].BoundingBox.Right;
                    if (gap > words[0].AvgFontSize * 0.5) // Significant gap
                    {
                        gaps.Add(gap);
                    }
                }

                // Check for regular column spacing
                if (gaps.Count >= 2)
                {
                    var avgGap = gaps.Average();
                    var stdDev = Math.Sqrt(gaps.Average(g => Math.Pow(g - avgGap, 2)));

                    // Low standard deviation indicates regular column spacing
                    if (stdDev < avgGap * 0.3)
                        return true;
                }
            }

            // Criterion 3: Short lines with multiple "columns" (delimited by 2+ spaces)
            if (line.BoundingBox.Width < pageWidth * 0.7)
            {
                var parts = Regex.Split(text, @"\s{2,}");
                if (parts.Length >= 3 && parts.All(p => p.Length < 50))
                    return true;
            }

            return false;
        }

        private static bool IsSignificantVerticalGap(TextLine currentLine, TextLine nextLine)
        {
            if (nextLine == null) return true;

            var currentBottom = currentLine.BoundingBox.Bottom;
            var nextTop = nextLine.BoundingBox.Top;
            var avgFontSize = (currentLine.AvgFontSize + nextLine.AvgFontSize) / 2;

            return (currentBottom - nextTop) > avgFontSize * 2.5;
        }

        private static DetectedTable AnalyzeTableStructure(List<TextLine> rows, double pageWidth)
        {
            var table = new DetectedTable();

            if (!rows.Any())
                return table;

            // Find column boundaries by analyzing word positions
            var allWords = rows.SelectMany(r => r.Words).ToList();
            var columnPositions = DetectColumnPositions(allWords, pageWidth);

            if (columnPositions.Count < 2)
            {
                // Fallback: create simple 2-column table
                columnPositions = new List<double> { 0, pageWidth / 2, pageWidth };
            }

            foreach (var row in rows)
            {
                var tableRow = new TableRow();
                var rowText = row.ToString().Trim();

                // Detect if it's a header (first row, often bold or capitalized)
                tableRow.IsHeader = rows.IndexOf(row) == 0 &&
                                  (row.Words.Any(w => w.IsBold) ||
                                   rowText.ToUpper() == rowText);

                // Split row into cells based on column positions
                for (int col = 0; col < columnPositions.Count - 1; col++)
                {
                    var cellStart = columnPositions[col];
                    var cellEnd = columnPositions[col + 1];
                    var cellMid = (cellStart + cellEnd) / 2;

                    // Find words that fall within this column
                    var cellWords = row.Words
                        .Where(w =>
                            (w.BoundingBox.Left >= cellStart && w.BoundingBox.Right <= cellEnd) ||
                            (w.BoundingBox.Left < cellEnd && w.BoundingBox.Right > cellStart) ||
                            Math.Abs(w.BoundingBox.Left + w.BoundingBox.Width / 2 - cellMid) < (cellEnd - cellStart) / 2)
                        .OrderBy(w => w.BoundingBox.Left)
                        .ToList();

                    var cellText = string.Join(" ", cellWords.Select(w => w.ToString())).Trim();

                    tableRow.Cells.Add(new TableCell
                    {
                        Text = cellText,
                        ColSpan = 1,
                        RowSpan = 1
                    });
                }

                // Merge empty cells horizontally
                MergeEmptyCells(tableRow);

                // Only add row if it has content
                if (tableRow.Cells.Any(c => !string.IsNullOrWhiteSpace(c.Text)))
                {
                    table.Rows.Add(tableRow);
                }
            }

            table.ColumnCount = columnPositions.Count - 1;

            // Set table bounds
            if (rows.Any())
            {
                var firstRow = rows.First();
                var lastRow = rows.Last();
                table.Bounds = new TextPosition
                {
                    X = firstRow.BoundingBox.Left,
                    Y = firstRow.BoundingBox.Top,
                    Width = firstRow.BoundingBox.Width,
                    Height = Math.Abs(firstRow.BoundingBox.Top - lastRow.BoundingBox.Bottom)
                };
            }

            return table;
        }

        #endregion

        #region Helper Methods

        private static List<TextLine> FilterOutTableLines(List<TextLine> lines, List<DetectedTable> tables)
        {
            if (!tables.Any())
                return lines;

            var tableLines = new HashSet<TextLine>();

            foreach (var table in tables)
            {
                foreach (var row in table.Rows)
                {
                    // Find the line that corresponds to this row
                    var rowText = string.Join(" ", row.Cells.Select(c => c.Text)).Trim();
                    var matchingLine = lines.FirstOrDefault(l =>
                        l.ToString().Trim().Equals(rowText, StringComparison.OrdinalIgnoreCase));

                    if (matchingLine != null)
                        tableLines.Add(matchingLine);
                }
            }

            return lines.Where(l => !tableLines.Contains(l)).ToList();
        }

        private static List<TextElement> GroupRelatedElements(List<TextElement> elements)
        {
            var grouped = new List<TextElement>();

            if (!elements.Any())
                return grouped;

            TextElement currentGroup = elements[0];

            for (int i = 1; i < elements.Count; i++)
            {
                var current = elements[i];
                var previous = elements[i - 1];

                // Check if current element should be merged with previous
                bool shouldMerge =
                    // Same type (except headings which should stay separate)
                    (current.Type == previous.Type &&
                     current.Type != ElementType.Heading1 &&
                     current.Type != ElementType.Heading2 &&
                     current.Type != ElementType.Heading3 &&
                     current.Type != ElementType.Heading4 &&
                     current.Type != ElementType.HorizontalRule &&
                     current.Type != ElementType.ListItem &&
                     current.Type != ElementType.NumberedListItem &&
                     current.Type != ElementType.Table) &&
                    // Similar formatting
                    Math.Abs(current.Format.FontSize - previous.Format.FontSize) < 0.5 &&
                    current.Format.IsBold == previous.Format.IsBold &&
                    current.Format.IsItalic == previous.Format.IsItalic &&
                    // Vertically close
                    Math.Abs(current.Position.Y - previous.Position.Y) < current.Format.FontSize * 2.5 &&
                    // Not significantly indented differently
                    Math.Abs(current.Position.X - previous.Position.X) < current.Format.FontSize * 2;

                if (shouldMerge)
                {
                    // Merge with previous
                    currentGroup.Text += " " + current.Text;
                    // Update confidence to the minimum
                    currentGroup.Confidence = Math.Min(currentGroup.Confidence, current.Confidence);
                }
                else
                {
                    grouped.Add(currentGroup);
                    currentGroup = current;
                }
            }

            // Add the last group
            grouped.Add(currentGroup);

            return grouped;
        }

        private static bool DetectUnderline(TextLine line)
        {
            // Simple heuristic: check if text is surrounded by underscores
            var text = line.ToString().Trim();
            return (text.StartsWith("_") && text.EndsWith("_") && text.Length > 2) ||
                   text.Contains("___") ||
                   text.StartsWith("<u>") || text.EndsWith("</u>");
        }

        private static int CalculateIndentLevel(double xPosition, double leftMargin)
        {
            if (xPosition <= leftMargin * 1.2)
                return 0;

            var indentUnits = (xPosition - leftMargin) / (leftMargin * 0.5);
            return Math.Max(0, (int)Math.Round(indentUnits));
        }

        private static List<double> DetectColumnPositions(List<TextWord> allWords, double pageWidth)
        {
            var columnPositions = new List<double>();

            if (!allWords.Any())
            {
                // Default columns
                columnPositions.Add(0);
                columnPositions.Add(pageWidth);
                return columnPositions;
            }

            // Get all X positions of word starts
            var xPositions = allWords
                .Select(w => w.BoundingBox.Left)
                .Where(x => x > 0 && x < pageWidth * 0.95)
                .OrderBy(x => x)
                .ToList();

            if (!xPositions.Any())
            {
                columnPositions.Add(0);
                columnPositions.Add(pageWidth);
                return columnPositions;
            }

            // Cluster X positions to find columns
            double clusterThreshold = pageWidth * 0.03; // 3% of page width
            var clusters = new List<List<double>>();
            var currentCluster = new List<double> { xPositions[0] };

            for (int i = 1; i < xPositions.Count; i++)
            {
                if (xPositions[i] - currentCluster.Last() < clusterThreshold)
                {
                    currentCluster.Add(xPositions[i]);
                }
                else
                {
                    if (currentCluster.Count > 1) // Only keep significant clusters
                    {
                        clusters.Add(new List<double>(currentCluster));
                    }
                    currentCluster = new List<double> { xPositions[i] };
                }
            }

            if (currentCluster.Count > 1)
            {
                clusters.Add(currentCluster);
            }

            // Get average X for each cluster
            var columnCenters = clusters.Select(c => c.Average()).OrderBy(x => x).ToList();

            // Add page boundaries
            columnPositions.Add(0);
            columnPositions.AddRange(columnCenters);
            columnPositions.Add(pageWidth);

            // Remove duplicates and sort
            columnPositions = columnPositions.Distinct().OrderBy(x => x).ToList();

            return columnPositions;
        }

        private static void MergeEmptyCells(TableRow row)
        {
            if (row.Cells.Count < 2)
                return;

            for (int i = row.Cells.Count - 1; i > 0; i--)
            {
                if (string.IsNullOrWhiteSpace(row.Cells[i].Text) &&
                    string.IsNullOrWhiteSpace(row.Cells[i - 1].Text))
                {
                    row.Cells[i - 1].ColSpan += row.Cells[i].ColSpan;
                    row.Cells.RemoveAt(i);
                }
            }
        }

        private static bool IsListItem(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return false;

            var trimmed = text.TrimStart();

            // Common list markers
            return trimmed.StartsWith("•") ||
                   trimmed.StartsWith("-") ||
                   trimmed.StartsWith("✓") ||
                   trimmed.StartsWith("▪") ||
                   trimmed.StartsWith("○") ||
                   trimmed.StartsWith("■") ||
                   trimmed.StartsWith("→") ||
                   trimmed.StartsWith("⇒") ||
                   Regex.IsMatch(trimmed, @"^[○◦◘◙▪•∙◉⦿◯○]\s") ||
                   Regex.IsMatch(trimmed, @"^[ivxIVX]+[\.\)]\s+") ||
                   Regex.IsMatch(trimmed, @"^[a-zA-Z][\.\)]\s+") ||
                   Regex.IsMatch(trimmed, @"^\d+[\.\)]\s+") ||
                   Regex.IsMatch(trimmed, @"^\(\d+\)\s+") ||
                   Regex.IsMatch(trimmed, @"^\[\d+\]\s+");
        }

        private static ElementType GetListItemType(string text)
        {
            var trimmed = text.TrimStart();

            if (Regex.IsMatch(trimmed, @"^\d+[\.\)]"))
                return ElementType.NumberedListItem;

            return ElementType.ListItem;
        }

        private static string CleanListItemText(string text)
        {
            var trimmed = text.TrimStart();

            // Patterns to remove (list markers)
            var patterns = new[]
            {
                @"^[•\-✓▪○■→⇒]\s*",
                @"^[○◦◘◙▪•∙◉⦿◯○]\s*",
                @"^[ivxIVX]+[\.\)]\s*",
                @"^[a-zA-Z][\.\)]\s*",
                @"^\d+[\.\)]\s*",
                @"^\(\d+\)\s*",
                @"^\[\d+\]\s*"
            };

            foreach (var pattern in patterns)
            {
                var match = Regex.Match(trimmed, pattern);
                if (match.Success)
                {
                    return trimmed.Substring(match.Length).Trim();
                }
            }

            return text.Trim();
        }

        private static bool IsMonospaceFont(string fontName)
        {
            if (string.IsNullOrEmpty(fontName))
                return false;

            var monoKeywords = new[]
            {
                "mono", "courier", "consolas", "terminal", "fixedsys",
                "source code", "dejavu sans mono", "liberation mono",
                "lucida console", "monaco", "andale mono", "roboto mono"
            };

            return monoKeywords.Any(keyword =>
                fontName.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private static bool IsHorizontalRule(string text)
        {
            var trimmed = text.Trim();

            if (trimmed.Length < 3)
                return false;

            // Check for repeating characters that form a line
            var firstChar = trimmed[0];
            if (firstChar != '-' && firstChar != '_' && firstChar != '*' && firstChar != '=' && firstChar != '~')
                return false;

            // At least 80% of the line should be the same character
            var sameCharCount = trimmed.Count(c => c == firstChar);
            return (double)sameCharCount / trimmed.Length > 0.8;
        }

        private static List<TextElement> DetectListNesting(List<TextElement> elements)
        {
            var result = new List<TextElement>();
            var listStack = new Stack<(int indent, ElementType type)>();

            // Track numbering for numbered lists
            int numberedListCounter = 1;
            ElementType lastNumberedType = ElementType.Unknown;

            foreach (var element in elements)
            {
                if (element.Type == ElementType.ListItem || element.Type == ElementType.NumberedListItem)
                {
                    var currentIndent = element.IndentLevel;

                    // Reset counter if starting new list or different indent level
                    if (listStack.Count == 0 ||
                        listStack.Peek().indent < currentIndent ||
                        (element.Type == ElementType.NumberedListItem && lastNumberedType != ElementType.NumberedListItem))
                    {
                        numberedListCounter = 1;
                    }

                    // Pop stack until we find parent level
                    while (listStack.Count > 0 && listStack.Peek().indent >= currentIndent)
                    {
                        listStack.Pop();
                    }

                    element.IndentLevel = listStack.Count;
                    listStack.Push((currentIndent, element.Type));
                    lastNumberedType = element.Type;
                }
                else
                {
                    // Reset list tracking when we leave list context
                    listStack.Clear();
                    numberedListCounter = 1;
                }

                result.Add(element);
            }

            return result;
        }

        private static string PostProcessMarkdown(string markdown)
        {
            if (string.IsNullOrEmpty(markdown))
                return markdown;

            var lines = markdown.Split('\n').ToList();
            var processed = new List<string>();

            // Track list context for proper numbering
            int listDepth = 0;
            var listCounters = new Dictionary<int, int>();

            for (int i = 0; i < lines.Count; i++)
            {
                var line = lines[i].TrimEnd();

                if (string.IsNullOrWhiteSpace(line))
                {
                    processed.Add(line);
                    continue;
                }

                // Detect list items and update numbering
                if (line.Contains("* ") || Regex.IsMatch(line, @"^\s*\d+\.\s"))
                {
                    // Count spaces at beginning to determine depth
                    var leadingSpaces = line.Length - line.TrimStart().Length;
                    var currentDepth = leadingSpaces / 2;

                    if (currentDepth > listDepth)
                    {
                        listDepth = currentDepth;
                        listCounters[listDepth] = 1;
                    }
                    else if (currentDepth < listDepth)
                    {
                        listDepth = currentDepth;
                    }

                    // Update numbered lists
                    if (Regex.IsMatch(line, @"^\s*\d+\.\s"))
                    {
                        if (!listCounters.ContainsKey(listDepth))
                            listCounters[listDepth] = 0;

                        listCounters[listDepth]++;
                        var counter = listCounters[listDepth];
                        line = Regex.Replace(line, @"^\s*\d+\.\s", new string(' ', leadingSpaces) + $"{counter}. ");
                    }
                }
                else
                {
                    // Reset list context when we leave lists
                    listDepth = 0;
                    listCounters.Clear();
                }

                // Merge with previous line if it's a continuation
                if (i > 0 && ShouldMergeLines(processed.Last(), line))
                {
                    var previous = processed.Last();
                    if (!previous.EndsWith(" ") && !line.StartsWith(" "))
                        processed[processed.Count - 1] = previous + " " + line;
                    else
                        processed[processed.Count - 1] = previous + line;
                }
                else
                {
                    processed.Add(line);
                }
            }

            // Clean up extra blank lines
            var cleaned = new List<string>();
            bool lastWasBlank = false;

            foreach (var line in processed)
            {
                if (string.IsNullOrWhiteSpace(line))
                {
                    if (!lastWasBlank)
                    {
                        cleaned.Add(line);
                        lastWasBlank = true;
                    }
                }
                else
                {
                    cleaned.Add(line);
                    lastWasBlank = false;
                }
            }

            return string.Join("\n", cleaned);
        }

        private static bool ShouldMergeLines(string previousLine, string currentLine)
        {
            if (string.IsNullOrWhiteSpace(previousLine) || string.IsNullOrWhiteSpace(currentLine))
                return false;

            var prevTrim = previousLine.TrimEnd();
            var currTrim = currentLine.TrimStart();

            // Don't merge if previous line ends with sentence-ending punctuation
            if (prevTrim.EndsWith(".") || prevTrim.EndsWith("!") || prevTrim.EndsWith("?") ||
                prevTrim.EndsWith(":") || prevTrim.EndsWith(";"))
                return false;

            // Don't merge headings, lists, or code blocks
            if (prevTrim.StartsWith("#") || prevTrim.StartsWith("*") || prevTrim.StartsWith("-") ||
                prevTrim.StartsWith(">") || prevTrim.StartsWith("```") || prevTrim.StartsWith("|") ||
                prevTrim.StartsWith("---") || prevTrim.StartsWith("==="))
                return false;

            // Don't merge if current line starts with uppercase (likely new sentence)
            if (currTrim.Length > 0 && char.IsUpper(currTrim[0]) &&
                !currTrim.StartsWith("I ") && !currTrim.StartsWith("I'"))
                return false;

            return true;
        }

        #endregion

        #region Markdown Formatting

        private static string FormatPageAsMarkdown(PageContent content, int pageNumber)
        {
            var mdBuilder = new StringBuilder();

            // Optional page header (can be commented out)
            // mdBuilder.AppendLine($"<!-- Page {pageNumber} -->\n");

            // Process non-table elements
            foreach (var element in content.Elements)
            {
                var elementMarkdown = FormatElementAsMarkdown(element);
                if (!string.IsNullOrWhiteSpace(elementMarkdown))
                    mdBuilder.Append(elementMarkdown);
            }

            // Process tables
            foreach (var table in content.Tables)
            {
                var tableMarkdown = FormatTableAsMarkdown(table);
                if (!string.IsNullOrWhiteSpace(tableMarkdown))
                {
                    mdBuilder.AppendLine(tableMarkdown);
                }
            }

            return mdBuilder.ToString();
        }

        private static string FormatElementAsMarkdown(TextElement element)
        {
            if (string.IsNullOrWhiteSpace(element.Text))
                return "";

            var text = element.Text;

            // Apply inline formatting (bold, italic, underline)
            if (element.Format.IsBold &&
                element.Type != ElementType.Heading1 &&
                element.Type != ElementType.Heading2 &&
                element.Type != ElementType.Heading3 &&
                element.Type != ElementType.Heading4)
            {
                // Check if already formatted
                if (!text.StartsWith("**") || !text.EndsWith("**"))
                    text = $"**{text}**";
            }

            if (element.Format.IsItalic)
            {
                // Avoid double formatting
                if (!text.StartsWith("*") || !text.EndsWith("*"))
                    text = $"*{text}*";
            }

            if (element.Format.IsUnderlined)
            {
                text = $"<u>{text}</u>";
            }

            // Apply element-specific formatting
            switch (element.Type)
            {
                case ElementType.Heading1:
                    return $"# {text}\n\n";

                case ElementType.Heading2:
                    return $"## {text}\n\n";

                case ElementType.Heading3:
                    return $"### {text}\n\n";

                case ElementType.Heading4:
                    return $"#### {text}\n\n";

                case ElementType.ListItem:
                    var indent = new string(' ', element.IndentLevel * 2);
                    return $"{indent}* {text}\n";

                case ElementType.NumberedListItem:
                    var numIndent = new string(' ', element.IndentLevel * 2);
                    // Number will be corrected in post-processing
                    return $"{numIndent}1. {text}\n";

                case ElementType.CodeBlock:
                    // Check if it's a single line or multi-line
                    if (text.Contains('\n') || text.Length > 60)
                        return $"```\n{text}\n```\n\n";
                    else
                        return $"`{text}`\n";

                case ElementType.BlockQuote:
                    var lines = text.Split('\n');
                    var quotedLines = lines.Select(l => $"> {l}");
                    return string.Join("\n", quotedLines) + "\n\n";

                case ElementType.HorizontalRule:
                    return "---\n\n";

                case ElementType.Footnote:
                    return $"[^{Guid.NewGuid().ToString("N").Substring(0, 4)}] {text}\n";

                case ElementType.PageNumber:
                    return ""; // Skip page numbers in output

                case ElementType.Paragraph:
                default:
                    return $"{text}\n\n";
            }
        }

        private static string FormatTableAsMarkdown(DetectedTable table)
        {
            if (table.Rows.Count == 0)
                return string.Empty;

            var mdBuilder = new StringBuilder();

            // Find maximum columns to ensure consistent formatting
            int maxColumns = table.Rows.Max(r => r.Cells.Count);

            // Pad rows with empty cells if needed
            foreach (var row in table.Rows)
            {
                while (row.Cells.Count < maxColumns)
                {
                    row.Cells.Add(new TableCell { Text = "" });
                }
            }

            // Format header row
            var headerRow = table.Rows.FirstOrDefault(r => r.IsHeader) ?? table.Rows.First();
            var headerCells = headerRow.Cells.Select(c => c.Text ?? "");
            mdBuilder.AppendLine("| " + string.Join(" | ", headerCells) + " |");

            // Add separator row
            var separators = headerRow.Cells.Select(c =>
            {
                var length = Math.Max(3, (c.Text ?? "").Length);
                return new string('-', Math.Min(length, 20)); // Limit to 20 chars
            });
            mdBuilder.AppendLine("| " + string.Join(" | ", separators) + " |");

            // Add data rows (skip header if it was the first row)
            var dataRows = table.Rows.Where(r => !r.IsHeader || table.Rows.IndexOf(r) > 0);
            foreach (var row in dataRows)
            {
                var cellTexts = row.Cells.Select(c => c.Text ?? "");
                mdBuilder.AppendLine("| " + string.Join(" | ", cellTexts) + " |");
            }

            mdBuilder.AppendLine(); // Add blank line after table

            return mdBuilder.ToString();
        }

        #endregion
    }
}
using BracketPipe;
using CsvHelper;
using CsvHelper.Configuration;
using DocSharp.Binary.DocFileFormat;
using DocSharp.Binary.OpenXmlLib;
using DocSharp.Binary.OpenXmlLib.PresentationML;
using DocSharp.Binary.OpenXmlLib.SpreadsheetML;
using DocSharp.Binary.OpenXmlLib.WordprocessingML;
using DocSharp.Binary.PptFileFormat;
using DocSharp.Binary.Spreadsheet.XlsFileFormat;
using DocSharp.Binary.StructuredStorage.Reader;
using DocSharp.Docx;
using MimeKit;
using RtfPipe;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;


namespace AllToMarkdown
{
    /// <summary>
    /// Static class providing methods to convert various file formats to Markdown.
    /// Supports text files, documents, spreadsheets, presentations, PDFs, and emails.
    /// </summary>
    public static class Converter
    {
        /// <summary>
        /// Supported file formats for conversion to Markdown
        /// </summary>
        public enum SupportedFileFormat
        {
            // Plain text and markup
            txt,
            md,

            // Word / RTF / OpenXML documents
            docx,
            doc,
            rtf,
            odt,

            // PDF
            pdf,

            // Spreadsheets
            xlsx,
            xls,
            csv,
            ods,

            // Presentations
            pptx,
            ppt,
            odp,

            // Email
            eml
        }

        /// <summary>
        /// Converts data from a stream to Markdown based on the specified file format.
        /// </summary>
        /// <param name="data">The stream containing the file data to convert.</param>
        /// <param name="fileFormat">The format of the data (e.g., PDF, DOCX, CSV).</param>
        /// <returns>A string containing the Markdown representation of the data.</returns>
        /// <exception cref="NotSupportedException">Thrown if the file format is not supported.</exception>
        public static string ConvertDataToMarkdown(Stream data, SupportedFileFormat fileFormat)
        {
            return fileFormat switch
            {
                SupportedFileFormat.md => new StreamReader(data).ReadToEnd(),
                SupportedFileFormat.txt => new StreamReader(data).ReadToEnd(),
                SupportedFileFormat.csv => ConvertCsvToMarkdown(data),
                SupportedFileFormat.docx => ConvertDocxToMarkdown(data),
                SupportedFileFormat.doc => ConvertDocBinaryToMarkdown(data),
                SupportedFileFormat.rtf => ConvertRtfToMarkdown(data),
                SupportedFileFormat.odt => ConvertOdtToMarkdown(data),
                SupportedFileFormat.xlsx => ConvertSpreadsheetToMarkdown(data, nameof(SupportedFileFormat.xlsx)),
                SupportedFileFormat.xls => ConvertSpreadsheetToMarkdown(data, nameof(SupportedFileFormat.xls)),
                SupportedFileFormat.ods => ConvertOdsToMarkdown(data),
                SupportedFileFormat.pptx => ConvertPresentationToMarkdown(data),
                SupportedFileFormat.ppt => ConvertPresentationToMarkdown(data),
                SupportedFileFormat.odp => ConvertOdpToMarkdown(data),
                SupportedFileFormat.pdf => ConvertPdf.ConvertPdfToMarkdown(data),
                SupportedFileFormat.eml => ConvertEmlToMarkdown(data),
                _ => throw new NotSupportedException($"Conversion for format '{fileFormat}' is not implemented.")
            };
        }
        

        /// <summary>
        /// Converts a file to Markdown by reading its content and determining the format from the file extension.
        /// </summary>
        /// <param name="filePath">The path to the file to convert.</param>
        /// <returns>A string containing the Markdown representation of the file content.</returns>
        /// <exception cref="ArgumentException">Thrown if the file path is null or whitespace.</exception>
        /// <exception cref="FileNotFoundException">Thrown if the file does not exist.</exception>
        /// <exception cref="NotSupportedException">Thrown if the file extension is not supported.</exception>
        public static string ConvertFileToMarkdown(string filePath)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(filePath, nameof(filePath));
            if (!File.Exists(filePath))
                throw new FileNotFoundException("File not found", filePath);

            var ext = Path.GetExtension(filePath);
            var supportedExtensions = Enum.GetNames(typeof(SupportedFileFormat)).Select(e => "." + e).ToArray();
            if (!supportedExtensions.Contains(ext, StringComparer.OrdinalIgnoreCase))
                throw new NotSupportedException($"Unsupported extension: {ext}");

            // Determine the file format from the extension
            var extWithoutDot = ext.TrimStart('.');
            if (!Enum.TryParse<SupportedFileFormat>(extWithoutDot, true, out var format))
                throw new NotSupportedException($"Unsupported extension: {ext}");

            using var stream = File.OpenRead(filePath);
            return ConvertDataToMarkdown(stream, format);
        }

        /// <summary>
        /// Converts DOCX data from a stream to Markdown using DocxToMarkdownConverter.
        /// Saves the stream to a temporary file for processing.
        /// </summary>
        private static string ConvertDocxToMarkdown(Stream data)
        {
            try
            {
                var tempFile = Path.GetTempFileName() + ".docx";
                using (var fileStream = File.Create(tempFile))
                {
                    data.CopyTo(fileStream);
                }

                try
                {
                    var converter = new DocxToMarkdownConverter()
                    {
                        ImagesOutputFolder = Path.GetTempPath(),
                        ImagesBaseUriOverride = string.Empty,
                        OriginalFolderPath = Path.GetTempPath()
                    };

                    var tempMd = Path.ChangeExtension(Path.GetTempFileName(), ".md");
                    converter.Convert(tempFile, tempMd);
                    var markdown = File.ReadAllText(tempMd);
                    File.Delete(tempMd);
                    return markdown;
                }
                finally
                {
                    if (File.Exists(tempFile))
                    {
                        File.Delete(tempFile);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to convert DOCX data: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Converts binary DOC data to Markdown by first converting to DOCX, then to Markdown.
        /// </summary>
        private static string ConvertDocBinaryToMarkdown(Stream data)
        {
            // Convert binary .doc to .docx using DocSharp.Binary, then to Markdown
            try
            {
                var tempDocPath = Path.GetTempFileName() + ".doc";
                using (var fileStream = File.Create(tempDocPath))
                {
                    data.CopyTo(fileStream);
                }

                try
                {
                    var tempDocxPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.docx");

                    try
                    {
                        // Step 1: Convert .doc binary to temporary .docx using DocSharp.Binary
                        using (var docReader = new StructuredStorageReader(tempDocPath))
                        {
                            var doc = new WordDocument(docReader);
                            using var docx = WordprocessingDocument.Create(tempDocxPath, WordprocessingDocumentType.Document);
                            DocSharp.Binary.WordprocessingMLMapping.Converter.Convert(doc, docx);
                        }

                        // Step 2: Convert .docx to Markdown using existing method
                        return ConvertDocxToMarkdown(File.OpenRead(tempDocxPath));
                    }
                    finally
                    {
                        // Clean up temporary file
                        if (File.Exists(tempDocxPath))
                        {
                            File.Delete(tempDocxPath);
                        }
                    }
                }
                finally
                {
                    if (File.Exists(tempDocPath))
                    {
                        File.Delete(tempDocPath);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to convert binary DOC data: {ex.Message}", ex);
            }
        }

        private static string ConvertRtfToMarkdown(Stream data)
        {
            // Extract text content from RTF file using RtfPipe and BracketPipe libraries
            try
            {
                var rtfContent = new StreamReader(data).ReadToEnd();

                // Step 1: Convert RTF to HTML using RtfPipe
                var htmlContent = Rtf.ToHtml(rtfContent);

                // Step 2: Convert HTML to Markdown using BracketPipe
                var markdown = Html.ToMarkdown(htmlContent);

                if (string.IsNullOrWhiteSpace(markdown))
                {
                    throw new InvalidOperationException("No readable text found in RTF file.");
                }

                return markdown;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to extract text from RTF data: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Parses CSV data and converts it to a Markdown table format.
        /// </summary>
        private static string ConvertCsvToMarkdown(Stream data)
        {
            // CSV files - use CsvHelper for robust parsing (handles quoted fields, escaped characters, etc.)
            try
            {
                var config = new CsvConfiguration(CultureInfo.InvariantCulture)
                {
                    HasHeaderRecord = true,
                    MissingFieldFound = null,
                    BadDataFound = null,
                    TrimOptions = TrimOptions.Trim
                };

                using var reader = new StreamReader(data);
                using var csv = new CsvReader(reader, config);

                var records = new List<string[]>();
                string[]? headerRow = null;

                // Read all records
                while (csv.Read())
                {
                    if (headerRow == null)
                    {
                        // First row is header
                        csv.ReadHeader();
                        headerRow = csv.HeaderRecord ?? [];
                        if (headerRow.Length > 0)
                        {
                            records.Add(headerRow);
                        }
                    }

                    // Read data row
                    var row = new List<string>();
                    for (int i = 0; i < (headerRow?.Length ?? csv.ColumnCount); i++)
                    {
                        row.Add(csv.TryGetField<string>(i, out var value) ? value ?? "" : "");
                    }
                    if (row.Count > 0)
                    {
                        records.Add(row.ToArray());
                    }
                }

                if (records.Count == 0)
                {
                    throw new InvalidOperationException("No data found in CSV file.");
                }

                return ConvertRecordsToMarkdownTable(records);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to convert CSV data: {ex.Message}", ex);
            }
        }

        private static string ConvertRecordsToMarkdownTable(List<string[]> records)
        {
            if (records.Count == 0)
                return string.Empty;

            var markdown = new System.Text.StringBuilder();
            var maxColumns = records.Max(r => r.Length);

            for (int i = 0; i < records.Count; i++)
            {
                var row = records[i];
                var cells = new List<string>();

                for (int j = 0; j < maxColumns; j++)
                {
                    var cell = j < row.Length ? row[j] : "";
                    // Escape pipe characters in cell content
                    cell = cell.Replace("|", "\\|");
                    cells.Add(cell);
                }

                markdown.AppendLine("| " + string.Join(" | ", cells) + " |");

                // Add header separator after first row
                if (i == 0)
                {
                    var separators = Enumerable.Repeat("---", maxColumns);
                    markdown.AppendLine("| " + string.Join(" | ", separators) + " |");
                }
            }

            return markdown.ToString();
        }

        private static string ConvertSpreadsheetToMarkdown(Stream data, string format)
        {
            // Convert Excel files using DocSharp.Binary for .xls, or extract from .xlsx
            try
            {
                if (format.Equals("xls", StringComparison.OrdinalIgnoreCase))
                {
                    // Convert binary .xls to .xlsx using DocSharp.Binary
                    var tempXlsPath = Path.GetTempFileName() + ".xls";
                    using (var fileStream = File.Create(tempXlsPath))
                    {
                        data.CopyTo(fileStream);
                    }

                    try
                    {
                        var tempXlsxPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");

                        try
                        {
                            using (var xlsReader = new StructuredStorageReader(tempXlsPath))
                            {
                                var xls = new XlsDocument(xlsReader);
                                using var xlsx = SpreadsheetDocument.Create(tempXlsxPath, SpreadsheetDocumentType.Workbook);
                                DocSharp.Binary.SpreadsheetMLMapping.Converter.Convert(xls, xlsx);
                            }

                            // Extract text from the .xlsx file
                            return ExtractTextFromXlsx(File.OpenRead(tempXlsxPath));
                        }
                        finally
                        {
                            if (File.Exists(tempXlsxPath))
                            {
                                File.Delete(tempXlsxPath);
                            }
                        }
                    }
                    finally
                    {
                        if (File.Exists(tempXlsPath))
                        {
                            File.Delete(tempXlsPath);
                        }
                    }
                }
                else
                {
                    // Handle .xlsx directly
                    return ExtractTextFromXlsx(data);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to extract text from {format} data: {ex.Message}", ex);
            }
        }

        private static string ExtractTextFromXlsx(Stream data)
        {
            // Extract text from .xlsx (OpenXML format - ZIP archive)
            try
            {
                using (var archive = new System.IO.Compression.ZipArchive(data, System.IO.Compression.ZipArchiveMode.Read))
                {
                    var sharedStringsEntry = archive.Entries.FirstOrDefault(e => e.FullName.EndsWith("sharedStrings.xml", StringComparison.OrdinalIgnoreCase));
                    var sheetEntries = archive.Entries.Where(e => e.FullName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase)).ToList();

                    var textBuilder = new System.Text.StringBuilder();

                    // Extract shared strings
                    if (sharedStringsEntry != null)
                    {
                        using (var stream = sharedStringsEntry.Open())
                        using (var reader = new StreamReader(stream))
                        {
                            var xmlContent = reader.ReadToEnd();
                            var text = ExtractTextFromXml(xmlContent);
                            if (!string.IsNullOrWhiteSpace(text))
                            {
                                textBuilder.AppendLine(text);
                            }
                        }
                    }

                    // Extract from sheets
                    foreach (var sheetEntry in sheetEntries)
                    {
                        using (var stream = sheetEntry.Open())
                        using (var reader = new StreamReader(stream))
                        {
                            var xmlContent = reader.ReadToEnd();
                            var text = ExtractTextFromXml(xmlContent);
                            if (!string.IsNullOrWhiteSpace(text))
                            {
                                textBuilder.AppendLine($"\n---\n{text}");
                            }
                        }
                    }

                    var result = textBuilder.ToString();
                    if (string.IsNullOrWhiteSpace(result))
                    {
                        throw new InvalidOperationException("No readable text found in XLSX file.");
                    }

                    return result;
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to extract text from XLSX data: {ex.Message}", ex);
            }
        }

        private static string ConvertOdtToMarkdown(Stream data)
        {
            // ODT (OpenDocument Text) - parse with ODF schema awareness
            try
            {
                return ExtractTextFromOdt(data);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to convert ODT data: {ex.Message}", ex);
            }
        }

        private static string ConvertOdsToMarkdown(Stream data)
        {
            // ODS (OpenDocument Spreadsheet) - extract tables
            try
            {
                return ExtractTextFromOds(data);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to convert ODS data: {ex.Message}", ex);
            }
        }

        private static string ConvertOdpToMarkdown(Stream data)
        {
            // ODP (OpenDocument Presentation) - extract slides with structure
            try
            {
                return ExtractTextFromOdp(data);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to convert ODP data: {ex.Message}", ex);
            }
        }

        private static string ExtractTextFromOdt(Stream data)
        {
            // Parse ODT content.xml with ODF namespace awareness
            using var archive = new System.IO.Compression.ZipArchive(data, System.IO.Compression.ZipArchiveMode.Read);

            var contentEntry = archive.Entries.FirstOrDefault(e => e.FullName == "content.xml");
            if (contentEntry == null)
                throw new FileNotFoundException("content.xml not found in ODT archive");

            using var stream = contentEntry.Open();
            using var reader = new StreamReader(stream);
            var xmlContent = reader.ReadToEnd();

            try
            {
                var doc = XDocument.Parse(xmlContent);

                // ODF namespaces
                XNamespace text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
                XNamespace table = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";

                var markdown = new System.Text.StringBuilder();

                // Process document body
                var body = doc.Descendants(text + "body").FirstOrDefault();
                if (body == null)
                    return ExtractTextFromXml(xmlContent);

                foreach (var element in body.Descendants())
                {
                    switch (element.Name.LocalName)
                    {
                        case "h": // Heading
                            var level = element.Attribute(text + "outline-level")?.Value ?? "1";
                            var headingPrefix = new string('#', int.Parse(level));
                            markdown.AppendLine($"{headingPrefix} {element.Value.Trim()}");
                            markdown.AppendLine();
                            break;
                        case "p": // Paragraph
                            var paraText = element.Value.Trim();
                            if (!string.IsNullOrWhiteSpace(paraText))
                            {
                                markdown.AppendLine(paraText);
                                markdown.AppendLine();
                            }
                            break;
                        case "list-item": // List item
                            markdown.AppendLine($"- {element.Value.Trim()}");
                            break;
                    }
                }

                return markdown.ToString();
            }
            catch
            {
                return ExtractTextFromXml(xmlContent);
            }
        }

        private static string ExtractTextFromOds(Stream data)
        {
            // Parse ODS content.xml to extract tables
            using var archive = new System.IO.Compression.ZipArchive(data, System.IO.Compression.ZipArchiveMode.Read);

            var contentEntry = archive.Entries.FirstOrDefault(e => e.FullName == "content.xml");
            if (contentEntry == null)
                throw new FileNotFoundException("content.xml not found in ODS archive");

            using var stream = contentEntry.Open();
            using var reader = new StreamReader(stream);
            var xmlContent = reader.ReadToEnd();

            try
            {
                var doc = XDocument.Parse(xmlContent);

                // ODF namespaces
                XNamespace table = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
                XNamespace text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

                var markdown = new System.Text.StringBuilder();
                var tables = doc.Descendants(table + "table").ToList();

                int tableNum = 1;
                foreach (var tbl in tables)
                {
                    var tableName = tbl.Attribute(table + "name")?.Value ?? $"Sheet {tableNum}";
                    markdown.AppendLine($"## {tableName}");
                    markdown.AppendLine();

                    var rows = tbl.Descendants(table + "table-row").ToList();
                    var tableData = new List<string[]>();

                    foreach (var row in rows)
                    {
                        var cells = row.Descendants(table + "table-cell").ToList();
                        var rowData = new List<string>();

                        foreach (var cell in cells)
                        {
                            var cellText = string.Join(" ", cell.Descendants(text + "p").Select(p => p.Value.Trim()));
                            rowData.Add(cellText);
                        }

                        if (rowData.Any(c => !string.IsNullOrWhiteSpace(c)))
                        {
                            tableData.Add(rowData.ToArray());
                        }
                    }

                    if (tableData.Count > 0)
                    {
                        markdown.AppendLine(ConvertRecordsToMarkdownTable(tableData));
                    }

                    markdown.AppendLine();
                    tableNum++;
                }

                return markdown.ToString();
            }
            catch
            {
                return ExtractTextFromXml(xmlContent);
            }
        }

        private static string ExtractTextFromOdp(Stream data)
        {
            // Parse ODP content.xml to extract slides
            using var archive = new System.IO.Compression.ZipArchive(data, System.IO.Compression.ZipArchiveMode.Read);

            var contentEntry = archive.Entries.FirstOrDefault(e => e.FullName == "content.xml");
            if (contentEntry == null)
                throw new FileNotFoundException("content.xml not found in ODP archive");

            using var stream = contentEntry.Open();
            using var reader = new StreamReader(stream);
            var xmlContent = reader.ReadToEnd();

            try
            {
                var doc = XDocument.Parse(xmlContent);

                // ODF namespaces
                XNamespace draw = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
                XNamespace text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
                XNamespace presentation = "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";

                var markdown = new System.Text.StringBuilder();
                var pages = doc.Descendants(draw + "page").ToList();

                int slideNum = 1;
                foreach (var page in pages)
                {
                    var slideName = page.Attribute(draw + "name")?.Value ?? $"Slide {slideNum}";
                    markdown.AppendLine($"## {slideName}");
                    markdown.AppendLine();

                    // Extract all text frames
                    var textFrames = page.Descendants(draw + "text-box")
                        .Concat(page.Descendants(draw + "frame"));

                    foreach (var frame in textFrames)
                    {
                        var paragraphs = frame.Descendants(text + "p");
                        foreach (var para in paragraphs)
                        {
                            var paraText = para.Value.Trim();
                            if (!string.IsNullOrWhiteSpace(paraText))
                            {
                                markdown.AppendLine(paraText);
                            }
                        }
                    }

                    // Extract notes
                    var notes = page.Descendants(presentation + "notes").FirstOrDefault();
                    if (notes != null)
                    {
                        var notesText = string.Join(" ", notes.Descendants(text + "p").Select(p => p.Value.Trim()));
                        if (!string.IsNullOrWhiteSpace(notesText))
                        {
                            markdown.AppendLine();
                            markdown.AppendLine("**Notes:**");
                            markdown.AppendLine(notesText);
                        }
                    }

                    markdown.AppendLine();
                    slideNum++;
                }

                return markdown.ToString();
            }
            catch
            {
                return ExtractTextFromXml(xmlContent);
            }
        }

        /// <summary>
        /// Extracts structured content from EML data and formats it as Markdown, including metadata and attachments.
        /// </summary>
        private static string ConvertEmlToMarkdown(Stream data)
        {
            // Extract structured content from EML (MIME) file using MimeKit
            // Handles complex MIME structures, encoding, attachments info
            try
            {
                var message = MimeMessage.Load(data);
                var markdown = new System.Text.StringBuilder();

                // Add Subject as main title
                if (!string.IsNullOrWhiteSpace(message.Subject))
                {
                    markdown.AppendLine($"# {message.Subject}");
                    markdown.AppendLine();
                }

                // Add metadata
                if (message.From.Count > 0)
                {
                    markdown.AppendLine($"**From:** {string.Join(", ", message.From)}");
                }

                if (message.To.Count > 0)
                {
                    markdown.AppendLine($"**To:** {string.Join(", ", message.To)}");
                }

                if (message.Cc.Count > 0)
                {
                    markdown.AppendLine($"**Cc:** {string.Join(", ", message.Cc)}");
                }

                if (message.Date != DateTimeOffset.MinValue)
                {
                    markdown.AppendLine($"**Date:** {message.Date:yyyy-MM-dd HH:mm:ss}");
                }

                markdown.AppendLine();
                markdown.AppendLine("---");
                markdown.AppendLine();

                // Extract body content
                var bodyText = ExtractEmailBodyFromMime(message);
                if (!string.IsNullOrWhiteSpace(bodyText))
                {
                    markdown.AppendLine(bodyText);
                }

                // List attachments
                var attachments = message.Attachments.ToList();
                if (attachments.Count > 0)
                {
                    markdown.AppendLine();
                    markdown.AppendLine("---");
                    markdown.AppendLine();
                    markdown.AppendLine("## Attachments");
                    foreach (var attachment in attachments)
                    {
                        var fileName = attachment.ContentDisposition?.FileName
                                 ?? (attachment as MimePart)?.FileName
                        ?? "Unknown";
                        var contentType = attachment.ContentType?.MimeType ?? "unknown";
                        markdown.AppendLine($"- {fileName} ({contentType})");
                    }
                }

                var result = markdown.ToString();

                if (string.IsNullOrWhiteSpace(result))
                {
                    throw new InvalidOperationException("No readable content found in EML file.");
                }

                return result;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to convert EML data: {ex.Message}", ex);
            }
        }

        private static string ExtractEmailBodyFromMime(MimeMessage message)
        {
            // Try to get text body first, then HTML converted to text
            var textBody = message.TextBody;
            if (!string.IsNullOrWhiteSpace(textBody))
            {
                return textBody.Trim();
            }

            // If only HTML body available, convert to markdown
            var htmlBody = message.HtmlBody;
            if (!string.IsNullOrWhiteSpace(htmlBody))
            {
                try
                {
                    // Use BracketPipe to convert HTML to Markdown
                    var markdown = Html.ToMarkdown(htmlBody);
                    return markdown?.Trim() ?? string.Empty;
                }
                catch
                {
                    // Fallback: strip HTML tags
                    return ExtractTextFromXml(htmlBody);
                }
            }

            // Try to extract from body parts recursively
            return ExtractTextFromMimePart(message.Body);
        }

        private static string ExtractTextFromMimePart(MimeEntity? entity)
        {
            if (entity == null) return string.Empty;

            switch (entity)
            {
                case TextPart textPart:
                    return textPart.Text?.Trim() ?? string.Empty;

                case Multipart multipart:
                    var textParts = new List<string>();
                    foreach (var part in multipart)
                    {
                        var text = ExtractTextFromMimePart(part);
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            textParts.Add(text);
                        }
                    }
                    return string.Join("\n\n", textParts);

                default:
                    return string.Empty;
            }
        }

        /// <summary>
        /// Converts presentation data (PPT or PPTX) to Markdown by extracting slide text.
        /// Detects format by attempting to open as ZIP archive.
        /// </summary>
        private static string ConvertPresentationToMarkdown(Stream data)
        {
            // Convert presentation files using DocSharp.Binary for .ppt, or extract from .pptx
            try
            {
                // Since we don't have extension here, assume based on content or handle both
                // For simplicity, try to detect or handle as pptx first, but since it's stream, we need to save temp anyway for ppt
                // Actually, since called from ConvertDataToMarkdown with specific format, but wait, the method is called for both ppt and pptx
                // In switch, it's pptx => ConvertPresentationToMarkdown(data), ppt => same
                // So, to distinguish, perhaps pass format, but for now, since ppt is binary, pptx is zip, we can try to open as zip, if fails, assume ppt
                // But complicated. Since the caller knows the format, perhaps modify to pass format.

                // For now, since pptx is zip, try ZipArchive, if succeeds, it's pptx, else save as ppt and convert.

                try
                {
                    using var archive = new System.IO.Compression.ZipArchive(data, System.IO.Compression.ZipArchiveMode.Read, leaveOpen: true);
                    // If succeeds, it's pptx
                    return ExtractTextFromPptx(data);
                }
                catch
                {
                    // Assume ppt, save temp .ppt
                    var tempPptPath = Path.GetTempFileName() + ".ppt";
                    using (var fileStream = File.Create(tempPptPath))
                    {
                        data.CopyTo(fileStream);
                    }

                    try
                    {
                        var tempPptxPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.pptx");

                        try
                        {
                            using (var pptReader = new StructuredStorageReader(tempPptPath))
                            {
                                var ppt = new PowerpointDocument(pptReader);
                                using var pptx = PresentationDocument.Create(tempPptxPath, PresentationDocumentType.Presentation);
                                DocSharp.Binary.PresentationMLMapping.Converter.Convert(ppt, pptx);
                            }

                            // Extract text from the .pptx file
                            return ExtractTextFromPptx(File.OpenRead(tempPptxPath));
                        }
                        finally
                        {
                            if (File.Exists(tempPptxPath))
                            {
                                File.Delete(tempPptxPath);
                            }
                        }
                    }
                    finally
                    {
                        if (File.Exists(tempPptPath))
                        {
                            File.Delete(tempPptPath);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to extract text from presentation data: {ex.Message}", ex);
            }
        }

        private static string ExtractTextFromPptx(Stream data)
        {
            // Extract text from .pptx (OpenXML format - ZIP archive)
            // Includes slide content, notes, and comments
            try
            {
                using var archive = new System.IO.Compression.ZipArchive(data, System.IO.Compression.ZipArchiveMode.Read);

                var slideEntries = archive.Entries
           .Where(e => e.FullName.StartsWith("ppt/slides/slide", StringComparison.OrdinalIgnoreCase)
                  && e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                         && !e.FullName.Contains("_rels"))
         .OrderBy(e => ExtractSlideNumber(e.FullName))
                    .ToList();

                var textBuilder = new System.Text.StringBuilder();
                int slideNumber = 1;

                foreach (var slideEntry in slideEntries)
                {
                    // Extract slide content
                    using (var stream = slideEntry.Open())
                    using (var reader = new StreamReader(stream))
                    {
                        var xmlContent = reader.ReadToEnd();
                        var text = ExtractTextFromPptxXml(xmlContent);

                        textBuilder.AppendLine($"## Slide {slideNumber}");
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            textBuilder.AppendLine(text);
                        }
                    }

                    // Extract speaker notes for this slide
                    var notesEntryName = $"ppt/notesSlides/notesSlide{slideNumber}.xml";
                    var notesEntry = archive.Entries.FirstOrDefault(e =>
                           e.FullName.Equals(notesEntryName, StringComparison.OrdinalIgnoreCase));

                    if (notesEntry != null)
                    {
                        using var notesStream = notesEntry.Open();
                        using var notesReader = new StreamReader(notesStream);
                        var notesXml = notesReader.ReadToEnd();
                        var notesText = ExtractTextFromPptxXml(notesXml);

                        if (!string.IsNullOrWhiteSpace(notesText))
                        {
                            textBuilder.AppendLine();
                            textBuilder.AppendLine("**Speaker Notes:**");
                            textBuilder.AppendLine(notesText);
                        }
                    }

                    textBuilder.AppendLine();
                    slideNumber++;
                }

                // Extract comments if present
                var commentsEntry = archive.Entries.FirstOrDefault(e =>
                  e.FullName.Contains("comments", StringComparison.OrdinalIgnoreCase)
               && e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase));

                if (commentsEntry != null)
                {
                    using var commentsStream = commentsEntry.Open();
                    using var commentsReader = new StreamReader(commentsStream);
                    var commentsXml = commentsReader.ReadToEnd();
                    var commentsText = ExtractTextFromPptxXml(commentsXml);

                    if (!string.IsNullOrWhiteSpace(commentsText))
                    {
                        textBuilder.AppendLine("---");
                        textBuilder.AppendLine("## Comments");
                        textBuilder.AppendLine(commentsText);
                    }
                }

                var result = textBuilder.ToString();
                if (string.IsNullOrWhiteSpace(result))
                {
                    throw new InvalidOperationException("No readable text found in PPTX file.");
                }

                return result;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to extract text from PPTX data: {ex.Message}", ex);
            }
        }

        private static int ExtractSlideNumber(string fullName)
        {
            // Extract slide number from path like "ppt/slides/slide1.xml"
            var match = Regex.Match(fullName, @"slide(\d+)\.xml", RegexOptions.IgnoreCase);
            return match.Success && int.TryParse(match.Groups[1].Value, out var num) ? num : 0;
        }

        private static string ExtractTextFromPptxXml(string xmlContent)
        {
            // Parse PPTX XML more intelligently to preserve text structure
            try
            {
                var doc = XDocument.Parse(xmlContent);

                // Define namespaces used in PPTX
                XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";

                var textParts = new List<string>();

                // Extract text from all <a:t> elements (text runs)
                var textElements = doc.Descendants(a + "t");
                foreach (var element in textElements)
                {
                    var text = element.Value?.Trim();
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        textParts.Add(text);
                    }
                }

                // Join with spaces, clean up multiple spaces
                var result = string.Join(" ", textParts);
                result = Regex.Replace(result, @"\s+", " ").Trim();

                return result;
            }
            catch
            {
                // Fallback to simple XML text extraction
                return ExtractTextFromXml(xmlContent);
            }
        }

        private static string ExtractTextFromXml(string xmlContent)
        {
            // Remove XML tags and extract plain text
            try
            {
                // Remove all XML tags
                var plainText = Regex.Replace(xmlContent, @"<[^>]+>", " ");
                // Remove XML entities
                plainText = System.Net.WebUtility.HtmlDecode(plainText);
                // Clean up whitespace
                plainText = Regex.Replace(plainText, @"\s+", " ");
                return plainText.Trim();
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string ExtractTextFromOpenDocument(Stream data, string entryName)
        {
            // OpenDocument files are ZIP archives
            // Extract and parse the XML content
            try
            {
                using (var archive = new System.IO.Compression.ZipArchive(data, System.IO.Compression.ZipArchiveMode.Read))
                {
                    var entry = archive.Entries.FirstOrDefault(e => e.Name == entryName || e.FullName.EndsWith(entryName));
                    if (entry == null)
                        throw new FileNotFoundException($"Entry {entryName} not found in archive");

                    using (var stream = entry.Open())
                    {
                        using (var reader = new StreamReader(stream))
                        {
                            var xmlContent = reader.ReadToEnd();
                            // Extract plain text from XML, removing tags
                            return ExtractTextFromXml(xmlContent);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to extract text from OpenDocument: {ex.Message}", ex);
            }
        }


    }
}

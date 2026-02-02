# AllToMarkdown

A comprehensive .NET library for converting various document formats to Markdown. Supports text files, Microsoft Office documents, OpenDocument formats, PDFs, and email files.

## Features

- **Wide Format Support**: Convert from TXT, MD, CSV, DOCX, DOC, RTF, ODT, XLSX, XLS, ODS, PPTX, PPT, ODP, PDF, and EML files to Markdown.
- **Structured Output**: Preserves document structure including headings, lists, tables, and formatting.
- **High Performance**: Optimized for processing large documents efficiently.
- **Easy Integration**: Simple API for .NET applications.
- **Cross-Platform**: Compatible with .NET 9 and later versions.

## Installation

Install the package via NuGet:

```
dotnet add package AllToMarkdown
```

Or download the source code and build the project.

## Usage

### Basic File Conversion

```csharp
using AllToMarkdown;

// Convert a file to Markdown
string markdown = Converter.ConvertFileToMarkdown("path/to/document.docx");
Console.WriteLine(markdown);
```

### Stream-Based Conversion

```csharp
using (var stream = File.OpenRead("path/to/document.pdf"))
{
    string markdown = Converter.ConvertDataToMarkdown(stream, SupportedFileFormat.pdf);
    Console.WriteLine(markdown);
}
```

### Supported Formats

- **Text Files**: TXT, MD
- **Spreadsheets**: CSV, XLSX, XLS, ODS
- **Word Documents**: DOCX, DOC, RTF, ODT
- **Presentations**: PPTX, PPT, ODP
- **PDFs**: PDF
- **Emails**: EML

## API Reference

### Converter Class

#### Methods

- `ConvertFileToMarkdown(string filePath)`: Converts a file at the specified path to Markdown.
- `ConvertDataToMarkdown(Stream data, SupportedFileFormat format)`: Converts data from a stream to Markdown based on the specified format.

#### SupportedFileFormat Enum

Defines the supported file formats for conversion.

## Requirements

- .NET 9.0 or later
- Dependencies: BracketPipe, CsvHelper, DocSharp libraries, MimeKit, PdfPig, RtfPipe

## Contributing

Contributions are welcome! Please fork the repository and submit pull requests for any improvements or bug fixes.

## Keywords

file converter, document to markdown, PDF to markdown, Office to markdown, .NET library, text processing, document parsing
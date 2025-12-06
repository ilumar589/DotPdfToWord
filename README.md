# DotPdfToWord

A simple .NET 10 console application that converts PDF documents to Word (DOCX) format.

## Features

- Convert PDF files to Word (DOCX) format
- Simple command-line interface
- Preserves basic text structure (paragraphs per page)
- Supports both absolute and relative file paths
- Built with .NET 10

## Requirements

- .NET 10 SDK or runtime

## Installation

### Build from source

```bash
git clone https://github.com/ilumar589/DotPdfToWord.git
cd DotPdfToWord
dotnet build -c Release
```

## Usage

```bash
DotPdfToWord <input.pdf> [output.docx]
```

### Arguments

- `<input.pdf>` - Path to the PDF file to convert (required)
- `[output.docx]` - Path to the output Word file (optional)
  - If not provided, the output will be saved as `<input>.docx`

### Examples

Convert a PDF to Word with automatic naming:
```bash
dotnet run --project DotPdfToWord document.pdf
# Creates document.docx
```

Convert a PDF to Word with custom output name:
```bash
dotnet run --project DotPdfToWord document.pdf my-document.docx
# Creates my-document.docx
```

Run the built executable:
```bash
./DotPdfToWord/bin/Release/net10.0/DotPdfToWord document.pdf output.docx
```

## Technology Stack

- **Framework**: .NET 10
- **Libraries**: 
  - PdfPig (UglyToad.PdfPig) - PDF text extraction
  - DocumentFormat.OpenXml - Word document generation (Microsoft's official SDK)

## Visual Studio

This solution can be opened directly in Visual Studio 2022:
1. Open `DotPdfToWord.sln` in Visual Studio
2. Build and run the project

## Notes

- The conversion extracts text content from PDFs and preserves basic structure
- Complex formatting, images, and layouts may not be preserved
- Each PDF page is converted to a paragraph in the Word document with page breaks

## License

This project is available for use under the terms specified in the repository.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

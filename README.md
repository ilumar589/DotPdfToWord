# DotPdfToWord

A simple .NET 10 console application that converts PDF documents to Word (DOCX) format.

## Features

- Convert PDF files to Word (DOCX) format
- Simple command-line interface
- Preserves text formatting and layout
- Supports both absolute and relative file paths
- Built with .NET 10

## Requirements

- .NET 10 SDK or runtime

## Installation

### Build from source

```bash
git clone https://github.com/ilumar589/DotPdfToWord.git
cd DotPdfToWord/DotPdfToWord
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
dotnet run document.pdf
# Creates document.docx
```

Convert a PDF to Word with custom output name:
```bash
dotnet run document.pdf my-document.docx
# Creates my-document.docx
```

Run the built executable:
```bash
./bin/Release/net10.0/DotPdfToWord document.pdf output.docx
```

## Technology Stack

- **Framework**: .NET 10
- **Libraries**: 
  - FreeSpire.PDF - PDF processing
  - FreeSpire.Doc - Word document generation

## Notes

- FreeSpire libraries have limitations in their free versions (e.g., page limits)
- For production use with large documents, consider purchasing a commercial license or using alternative libraries
- The conversion quality depends on the complexity of the PDF document

## License

This project is available for use under the terms specified in the repository.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

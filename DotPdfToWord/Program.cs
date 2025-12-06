using UglyToad.PdfPig;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

// Check command-line arguments
if (args.Length < 1)
{
    Console.WriteLine("PDF to Word Converter");
    Console.WriteLine("=====================");
    Console.WriteLine();
    Console.WriteLine("Usage: DotPdfToWord <input.pdf> [output.docx]");
    Console.WriteLine();
    Console.WriteLine("Arguments:");
    Console.WriteLine("  <input.pdf>   - Path to the PDF file to convert (required)");
    Console.WriteLine("  [output.docx] - Path to the output Word file (optional)");
    Console.WriteLine("                  If not provided, the output will be saved as <input>.docx");
    Console.WriteLine();
    Console.WriteLine("Example:");
    Console.WriteLine("  DotPdfToWord document.pdf");
    Console.WriteLine("  DotPdfToWord document.pdf output.docx");
    return 1;
}

string inputPath = args[0];
string outputPath;

// Determine output path
if (args.Length >= 2)
{
    outputPath = args[1];
}
else
{
    // Use input filename with .docx extension
    outputPath = Path.ChangeExtension(inputPath, ".docx");
}

try
{
    // Validate input file exists
    if (!File.Exists(inputPath))
    {
        Console.Error.WriteLine($"Error: Input file '{inputPath}' not found.");
        return 1;
    }

    // Validate input file is a PDF
    if (!Path.GetExtension(inputPath).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
    {
        Console.Error.WriteLine($"Error: Input file must be a PDF file.");
        return 1;
    }

    Console.WriteLine($"Converting '{inputPath}' to '{outputPath}'...");

    // Extract text from PDF using PdfPig
    using (var pdfDocument = PdfDocument.Open(inputPath))
    {
        // Create a new Word document using OpenXml
        using (var wordDocument = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document))
        {
            // Add main document part
            var mainPart = wordDocument.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = mainPart.Document.AppendChild(new Body());

            // Process each page of the PDF
            foreach (var page in pdfDocument.GetPages())
            {
                // Extract text from the page
                var pageText = page.Text;

                // Add a paragraph for this page
                var paragraph = body.AppendChild(new Paragraph());
                var run = paragraph.AppendChild(new Run());
                run.AppendChild(new Text(pageText));

                // Add a page break after each page (except the last one)
                if (page.Number < pdfDocument.NumberOfPages)
                {
                    var breakParagraph = body.AppendChild(new Paragraph());
                    var breakRun = breakParagraph.AppendChild(new Run());
                    breakRun.AppendChild(new Break() { Type = BreakValues.Page });
                }
            }

            // Save the document
            mainPart.Document.Save();
        }
    }

    Console.WriteLine("Conversion completed successfully!");
    Console.WriteLine($"Output saved to: {outputPath}");
    return 0;
}
catch (Exception ex)
{
    Console.Error.WriteLine($"Error during conversion: {ex.Message}");
    return 1;
}

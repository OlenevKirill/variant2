using System;
using System.IO;
using System.Text;
using DocX = Xceed.Words.NET.DocX;
using OfficeOpenXml;
using DinkToPdf;
using DinkToPdf.Contracts;

class Program
{
    static void Main(string[] args)
    {
        ConvertDocxToHtml("example.docx", "output.html");
        ConvertHtmlToPdf("example.html", "output.pdf");
    }

    static void ConvertDocxToHtml(string inputPath, string outputPath)
    {
        using (var document = DocX.Load(inputPath))
        {
            var html = document.Text; 
            File.WriteAllText(outputPath, html);
        }
    }
       

    static void ConvertHtmlToPdf(string htmlPath, string pdfPath)
    {
        var converter = new BasicConverter(new PdfTools());
        var doc = new HtmlToPdfDocument()
        {
            GlobalSettings = { ColorMode = ColorMode.Color, Orientation = Orientation.Portrait, PaperSize = PaperKind.A4 },
            Objects = {new ObjectSettings { HtmlContent = File.ReadAllText(htmlPath)}}
        };

        var pdf = converter.Convert(doc);

        using (var stream = new FileStream(pdfPath, FileMode.Create))
        {
            stream.Write(pdf, 0, pdf.Length);
        }
    }
}


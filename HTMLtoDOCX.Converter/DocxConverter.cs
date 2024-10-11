using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using HtmlToOpenXml;

namespace HTMLtoDOCX.Converter;

public static class DocxConverter
{
    public static async Task<Exception?> ConvertToDocx(string html, string filename, CancellationToken token = default)
    {
        try
        {
            if (File.Exists(filename))
            {
                File.Delete(filename);
            }

            using MemoryStream generatedDocument = new MemoryStream();
            using WordprocessingDocument package = WordprocessingDocument.Create(generatedDocument, WordprocessingDocumentType.Document);
            
            MainDocumentPart? mainPart = package.MainDocumentPart;
            
            if (mainPart is null)
            {
                mainPart = package.AddMainDocumentPart();
                new Document(new Body()).Save(mainPart);
            }

            HtmlConverter converter = new HtmlConverter(mainPart);
            await converter.ParseBody(html, token);

            mainPart.Document.Save();

            await File.WriteAllBytesAsync(filename, generatedDocument.ToArray(), token);
            return null;
        }
        catch (Exception e)
        {
            return e;
        }
    }

}
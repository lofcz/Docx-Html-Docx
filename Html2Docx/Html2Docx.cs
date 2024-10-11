using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml;

namespace Html2DocxCore;

/// <summary>
/// Converts HTML to Docx
/// </summary>
public static class Html2Docx
{
    /// <summary>
    /// Converts given HTML to DOCX
    /// </summary>
    /// <param name="html">HTML to convert</param>
    /// <param name="filename">Path where the output will be stored. Existing file will be deleted if necessary</param>
    /// <param name="headerHtml">Header HTML, will be displayed on every page</param>
    /// <param name="footerHtml">Footer HTML, will be displayed on every page</param>
    /// <param name="token">Optional cancellation token</param>
    /// <returns></returns>
    public static async Task<Exception?> Convert(string html, string filename, string? headerHtml = null, string? footerHtml = null, CancellationToken token = default)
    {
        try
        {
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

            if (headerHtml is not null)
            {
                await converter.ParseHeader(headerHtml, HeaderFooterValues.Default, token);
            }

            if (footerHtml is not null)
            {
                await converter.ParseFooter(footerHtml, HeaderFooterValues.Default, token);
            }
            
            mainPart.Document.Save();

            if (File.Exists(filename))
            {
                File.Delete(filename);
            }
            
            await File.WriteAllBytesAsync(filename, generatedDocument.ToArray(), token);
            return null;
        }
        catch (Exception e)
        {
            return e;
        }
    }
}
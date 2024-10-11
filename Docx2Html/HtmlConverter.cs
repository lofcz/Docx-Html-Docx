using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;

namespace Docx2Html;

public static class HtmlConverter
{
    public static async Task<string> ConvertToHtml(string filePath)
    {
        if (string.IsNullOrEmpty(filePath) || Path.GetExtension(filePath) != ".docx")
        {
            return "Unsupported format";
        }

        FileInfo fileInfo = new FileInfo(filePath);

        string htmlText = string.Empty;
        try
        {
            htmlText = await ParseDocx(fileInfo);
        }
        catch (OpenXmlPackageException e)
        {
            if (e.ToString().Contains("Invalid Hyperlink"))
            {
                await using (FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                {
                    UriFixer.FixInvalidUri(fs, FixUri);
                }
                
                htmlText = await ParseDocx(fileInfo);
            }
        }
        
        return htmlText;
    }

    private static string FixUri(string brokenUri)
    {
        string newUri;

        if (brokenUri.Contains("mailto:"))
        {
            int mailToCount = "mailto:".Length;
            brokenUri = brokenUri.Remove(0, mailToCount);
            newUri = brokenUri;
        }
        else
        {
            newUri = " ";
        }
        return newUri;
    }
    
    private static async Task<string> ParseDocx(FileInfo fileInfo)
    {
        try
        {
            byte[] byteArray = await File.ReadAllBytesAsync(fileInfo.FullName);
            using MemoryStream memoryStream = new MemoryStream();
            memoryStream.Write(byteArray, 0, byteArray.Length);

            using WordprocessingDocument wDoc = WordprocessingDocument.Open(memoryStream, true);

            string pageTitle = fileInfo.FullName;
            CoreFilePropertiesPart? part = wDoc.CoreFilePropertiesPart;
            
            if (part is not null)
            {
                pageTitle = (string?)part.GetXDocument().Descendants(DC.title).FirstOrDefault() ?? fileInfo.FullName;
            }

            WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings
            {
                AdditionalCss = "body { margin: 1cm auto; max-width: 20cm; padding: 0; }",
                PageTitle = pageTitle,
                FabricateCssClasses = true,
                CssClassPrefix = "pt-",
                RestrictToSupportedLanguages = false,
                RestrictToSupportedNumberingFormats = false,
                ImageHandler = imageInfo =>
                {
                    string extension = imageInfo.ContentType.Split('/')[1].ToLower();
                    
                    ImageFormat? imageFormat = extension switch
                    {
                        "png" => ImageFormat.Png,
                        "gif" => ImageFormat.Gif,
                        "bmp" => ImageFormat.Bmp,
                        "jpeg" => ImageFormat.Jpeg,
                        "tiff" => ImageFormat.Gif,
                        "x-wmf" => ImageFormat.Wmf,
                        _ => null
                    };

                    if (imageFormat == null)
                        return null;

                    string? base64;
                    
                    try
                    {
                        using MemoryStream ms = new MemoryStream();
                        imageInfo.Bitmap.Save(ms, imageFormat);
                        byte[] ba = ms.ToArray();
                        base64 = System.Convert.ToBase64String(ba);
                    }
                    catch (System.Runtime.InteropServices.ExternalException)
                    {
                        return null;
                    }

                    ImageFormat format = imageInfo.Bitmap.RawFormat;
                    ImageCodecInfo codec = ImageCodecInfo.GetImageDecoders().First(c => c.FormatID == format.Guid);
                    string? mimeType = codec.MimeType;

                    string imageSource = $"data:{mimeType};base64,{base64}";

                    XElement img = new XElement(Xhtml.img, new XAttribute(NoNamespace.src, imageSource), imageInfo.ImgStyleAttribute, imageInfo.AltText != null ? new XAttribute(NoNamespace.alt, imageInfo.AltText) : null);
                    return img;
                }
            };

            XElement? htmlElement = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);

            XDocument html = new XDocument(new XDocumentType("html", null, null, null), htmlElement);
            string htmlString = html.ToString(SaveOptions.DisableFormatting);
            return htmlString;
        }
        catch
        {
            return "File contains corrupt data";
        }
    }
}
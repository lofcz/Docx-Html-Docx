using System.IO;
using System.Linq;
using OpenXmlPowerTools;
using DocumentFormat.OpenXml.Packaging;
using System.Drawing.Imaging;
using System.Xml.Linq;

namespace DocxToHTML.Converter;

public class HTMLConverter
{


    public string ConvertToHtml(string fullFilePath)
    {
        if (string.IsNullOrEmpty(fullFilePath) || Path.GetExtension(fullFilePath) != ".docx")
            return "Unsupported format";

        var fileInfo = new FileInfo(fullFilePath);

        var htmlText = string.Empty;
        try
        {
            htmlText = ParseDOCX(fileInfo);
        }
        catch (OpenXmlPackageException e)
        {

            if (e.ToString().Contains("Invalid Hyperlink"))
            {
                using (var fs = new FileStream(fullFilePath, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                {
                    UriFixer.FixInvalidUri(fs, FixUri);
                }
                htmlText = ParseDOCX(fileInfo);
            }
        }


        return htmlText;


    }

    private static string FixUri(string brokenUri)
    {
        string newUri;

        if (brokenUri.Contains("mailto:"))
        {
            var mailToCount = "mailto:".Length;
            brokenUri = brokenUri.Remove(0, mailToCount);
            newUri = brokenUri;
        }
        else
        {
            newUri = " ";
        }
        return newUri;
    }


    private string ParseDOCX(FileInfo fileInfo)
    {

        try
        {
            var byteArray = File.ReadAllBytes(fileInfo.FullName);
            using var memoryStream = new MemoryStream();
            memoryStream.Write(byteArray, 0, byteArray.Length);

            using var wDoc = WordprocessingDocument.Open(memoryStream, true);

            var pageTitle = fileInfo.FullName;
            var part = wDoc.CoreFilePropertiesPart;
            if (part != null)
                pageTitle = (string)part.GetXDocument().Descendants(DC.title).FirstOrDefault() ?? fileInfo.FullName;

            var settings = new WmlToHtmlConverterSettings()
            {
                AdditionalCss = "body { margin: 1cm auto; max-width: 20cm; padding: 0; }",
                PageTitle = pageTitle,
                FabricateCssClasses = true,
                CssClassPrefix = "pt-",
                RestrictToSupportedLanguages = false,
                RestrictToSupportedNumberingFormats = false,
                ImageHandler = imageInfo =>
                {
                    var extension = imageInfo.ContentType.Split('/')[1].ToLower();
                    ImageFormat? imageFormat = null;
                    if (extension == "png") imageFormat = ImageFormat.Png;
                    else if (extension == "gif") imageFormat = ImageFormat.Gif;
                    else if (extension == "bmp") imageFormat = ImageFormat.Bmp;
                    else if (extension == "jpeg") imageFormat = ImageFormat.Jpeg;
                    else if (extension == "tiff")
                    {
                        imageFormat = ImageFormat.Gif;
                    }
                    else if (extension == "x-wmf")
                    {
                        imageFormat = ImageFormat.Wmf;
                    }

                    if (imageFormat == null)
                        return null;

                    string? base64;
                    try
                    {
                        using var ms = new MemoryStream();
                        imageInfo.Bitmap.Save(ms, imageFormat);
                        var ba = ms.ToArray();
                        base64 = System.Convert.ToBase64String(ba);
                    }
                    catch (System.Runtime.InteropServices.ExternalException)
                    { return null; }


                    var format = imageInfo.Bitmap.RawFormat;
                    var codec = ImageCodecInfo.GetImageDecoders().First(c => c.FormatID == format.Guid);
                    var mimeType = codec.MimeType;

                    var imageSource = string.Format("data:{0};base64,{1}", mimeType, base64);

                    var img = new XElement(Xhtml.img,
                        new XAttribute(NoNamespace.src, imageSource),
                        imageInfo.ImgStyleAttribute,
                        imageInfo.AltText != null ?
                            new XAttribute(NoNamespace.alt, imageInfo.AltText) : null);
                    return img;
                }
            };

            var htmlElement = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);

            var html = new XDocument(new XDocumentType("html", null, null, null), htmlElement);
            var htmlString = html.ToString(SaveOptions.DisableFormatting);
            return htmlString;
        }
        catch
        {
            return "File contains corrupt data";
        }

    }


}
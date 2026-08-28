using System.Diagnostics;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Contoso.Portal.Data.DAL.Helpers;

namespace Contoso.Portal.Common
{

    public class DocumentFromTemplateGenerator
    {

        /// <summary>
        /// Generate a .docx file from template, and fill it with data in XML.
        /// </summary>
        /// <param name="streamTemplate">Stream with template.</param>
        /// <param name="xmlData">XML with data to fill template.</param>
        /// <returns>TemporalFile create</returns>
        public static async Task<byte[]> GenerateDocumentFromTemplateAndXML(Stream streamTemplate, XDocument xmlData)
        {
            using var wordDoc = WordprocessingDocument.Open(streamTemplate, true);
            try
            {
                var mainPart = wordDoc.MainDocumentPart!; // Get the main part of the document which contains CustomXMLParts.
                mainPart.DeleteParts(mainPart.CustomXmlParts); // Delete all CustomXMLParts in the document. If needed only specific CustomXMLParts can be deleted using the CustomXmlParts IEnumerable.
                var myXmlPart = mainPart.AddCustomXmlPart(CustomXmlPartType.CustomXml); // Add new CustomXMLPart with data from new XML file.

                using var stream = new MemoryStream();
                await xmlData.SaveAsync(stream, SaveOptions.None, CancellationToken.None);
                stream.Position = 0;
                myXmlPart.FeedData(stream);

                using var response = new MemoryStream();
                wordDoc.Clone(response, false);
                wordDoc.Dispose();
                return response.ToArray();
            }
            catch (Exception ex)
            {
                wordDoc.Dispose();
                throw new Exception($"Error {nameof(GenerateDocumentFromTemplateAndXML)} with: xml='{xmlData.ToString()}' template size='{streamTemplate.Length}'", ex);
            }
        }

    }
} // namespace
using System;
using Newtonsoft.Json.Linq;
using System.ComponentModel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Frends.Office.Standard
{
    /// <summary>
    /// Used for writing Excel files.
    /// </summary>
    public class WriteWordFile
    {
        /// <summary>
        /// Allows you to write word files in .docx format. Repository: https://github.com/MarcinMichnik-HiQ/Frends.Office.Standard
        /// </summary>
        /// <param name="input"></param>
        /// <returns>Returns JToken.</returns>
        public static JToken WriteWordFileTask([PropertyTab] WriteWordFileInput input)
        {
            JToken taskResponse = JToken.Parse("{}");
            try
            {
                using (WordprocessingDocument wordDocument =
                    WordprocessingDocument.Create(input.TargetPath, WordprocessingDocumentType.Document))
                {
                    // Add a main document part. 
                    MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                    // Create the document structure and add some text.
                    mainPart.Document = new Document();
                    Body body = mainPart.Document.AppendChild(new Body());
                    Paragraph para = body.AppendChild(new Paragraph());
                    Run run = para.AppendChild(new Run());

                    string[] records = input.StringInput.Split(new string[] { input.LineDelimiter }, StringSplitOptions.None);

                    foreach (string record in records)
                    {
                        run.AppendChild(new Text(record));
                        run.Append(new Break());
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Unable to build and save word file.", ex);
            }

            taskResponse["message"] = "The file has been written correctly.";
            taskResponse["savedTo"] = input.TargetPath;

            return taskResponse;
        }
    }
}

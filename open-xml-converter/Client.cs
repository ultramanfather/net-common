using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;
using OfficeConverter;

namespace OpenXmlConverter
{
    public static class Client
    {
        public static void HTML2Word(Library.Model.HTML2WordParam param)
        {
            Console.WriteLine(@"Start:{0}", DateTime.Now.ToString());
            double columnWidth = param.Setting.PageWidth - param.Setting.MarginLeft - param.Setting.MarginRight;
            if (param.Setting.Columns > 1)
            {
                columnWidth = columnWidth / param.Setting.Columns;
            }

            using (WordprocessingDocument package =
                WordprocessingDocument.Create(param.Filepath, WordprocessingDocumentType.Document, true))
            {
                MainDocumentPart mainPart = package.MainDocumentPart;
                if (mainPart == null)
                {
                    mainPart = package.AddMainDocumentPart();
                    new Document(new Body()).Save(mainPart);
                }

                HtmlConverter converter = new HtmlConverter(mainPart, columnWidth, param.Setting);

                converter.ParseHtml(param.HTML);

                CustomSetting cs = new CustomSetting(param.Setting, mainPart);
                cs.Setting();
                mainPart.Document.Save();

                // AssertThatOpenXmlDocumentIsValid(package); // TODO
            }
            if (param.Setting.Convert2PDF == "1")
            {
                ExportOperating operating = new ExportOperating(param.Filepath, param.PDFPath, param.Setting);
                operating.Export();
            }
            Console.WriteLine(@"Over:{0}", DateTime.Now.ToString());
        }


        static void AssertThatOpenXmlDocumentIsValid(WordprocessingDocument wpDoc)
        {
            var validator = new OpenXmlValidator(FileFormatVersions.Office2010);
            var errors = validator.Validate(wpDoc);

            if (!errors.GetEnumerator().MoveNext())
                return;

            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("The document doesn't look 100% compatible with Office 2010.\n");

            Console.ForegroundColor = ConsoleColor.Gray;
            foreach (ValidationErrorInfo error in errors)
            {
                Console.Write("{0}\n\t{1}", error.Path.XPath, error.Description);
                Console.WriteLine();
            }

            Console.ReadLine();
        }
    }
}
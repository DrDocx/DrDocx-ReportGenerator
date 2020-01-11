using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ReportGenCLI
{
    static class Program
    {
        static void Main(string[] args)
        {
            PatientData patient = new PatientData("Joe Doe", "Johnny", "02/17/1988");
            string templatePath = @"Reports\ReportTemplate\Report_Template.dotx";
            string newfilePath = @"Reports\GeneratedReports\" + patient.name + ".docx";

            if (File.Exists(newfilePath))
            {
                File.Delete(newfilePath);
            }

            File.Copy(templatePath, newfilePath);

            byte[] byteArray = File.ReadAllBytes(newfilePath);

            using (var stream = new MemoryStream())
            {
                stream.Write(byteArray, 0, byteArray.Length);

                using (WordprocessingDocument doc = WordprocessingDocument.Open(stream, true))
                {
                    //Needed because I'm working with template dotx file, 
                    //remove this if the template is a normal docx. 
                    doc.ChangeDocumentType(DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
                    doc.InsertText("NAME", "testtesttesttest");
                }
                using (FileStream fs = new FileStream(newfilePath, FileMode.Create))
                {
                    stream.WriteTo(fs);
                }
            }

            //Console.WriteLine(makeRegex("name"));
            //SearchAndReplace(newfilePath,makeRegex("name"),patient.name);
            //SearchAndReplace(newfilePath,"_preferredName",patient.preferredName);

            /*
            Document doc = new Document();
            doc.LoadFromFile(templatePath);
            doc.Replace("fname", patient.name, true, true);
            doc.Replace("fpreferredName",patient.preferredName,true,true);
            doc.Replace("fdob",patient.dob,true,true);
            doc.Replace("_age","43",true,true);
            doc.SaveToFile(newfilePath, FileFormat.Docx2013);
            //SearchAndReplace(newfilePath,"Evaluation Warning: The document was created with Spire.Doc for .NET.","");
            */

            Console.WriteLine("Modified");

        }

        public static WordprocessingDocument InsertText(this WordprocessingDocument doc, string contentControlTag, string text)
        {
            var filteredBodyContentControls = doc.MainDocumentPart.Document.Body.Descendants<SdtElement>()
            .Where(sdt => sdt.SdtProperties.GetFirstChild<Tag>()?.Val == contentControlTag);

            var header = doc.MainDocumentPart.HeaderParts;
            foreach (var headerPart in header)
            {
                var headerContentControls = headerPart.Header.Descendants<SdtElement>();
                var filteredCCs = headerContentControls.Where(sdt => sdt.SdtProperties.GetFirstChild<Tag>()?.Val == contentControlTag);
                foreach (var contentControl in filteredCCs)
                {
                    contentControl.Descendants<Text>().First().Text = text;
                }
            }

            var footer = doc.MainDocumentPart.FooterParts;
            foreach (var footerPart in footer)
            {
                var footerContentControls = footerPart.Footer.Descendants<SdtElement>();
                var filteredCCs = footerContentControls.Where(sdt => sdt.SdtProperties.GetFirstChild<Tag>()?.Val == contentControlTag);
                foreach (var contentControl in filteredCCs)
                {
                    contentControl.Descendants<Text>().First().Text = text;
                }
            }

            foreach (var contentControl in filteredBodyContentControls)
            {
                contentControl.Descendants<Text>().First().Text = text;
            }

            //element.Descendants<Text>().First().Text = text;
            //element.Descendants<Text>().Skip(1).ToList().ForEach(t => t.Remove());

            return doc;
        }
    }
    public struct PatientData
    {
        public string name;
        public string preferredName;
        public string dob;
        public PatientData(string Name, string PrefName, string Dob)
        {
            name = Name;
            preferredName = PrefName;
            dob = Dob;
        }
    }
}

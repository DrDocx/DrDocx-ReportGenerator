using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

using Models;
using WordDocEditing;

namespace ReportGenCLI
{
	static class Program
	{
		static void Main(string[] args)
		{

			Random random = new Random();
			List<TestResult> results = new List<TestResult>();
			for (int i = 0; i < 8; i++)
			{
				double r = random.NextDouble() * 6 - 3;
				results.Add(new TestResult(){
					RelatedTest = new Test(){Name = "Kidney Digestion Evaluation"},
					ZScore = (int)(random.NextDouble()*6 - 3),
					Percentile = (int)(random.NextDouble()*100)
					});
			}

			List<TestResultGroup> resultGroups = new List<TestResultGroup>(new TestResultGroup[]{
				new TestResultGroup(){
					TestGroupInfo = new TestGroup(){Name = "Symptom Checklist - 90 - Revised"},
					Tests = results
				}
				});

			Patient johnDoe = new Patient(){
				Name = "John Doe",
				PreferredName = "Johnny",
				DateOfBirth = new DateTime(628243894905389400),
				DateOfTesting = new DateTime(628243894905389400),
				MedicalRecordNumber = 123456,
				ResultGroups = resultGroups
			};

			Patient patient = johnDoe;

			Console.WriteLine(patient.DateOfBirth);
			

			string templatePath = @"Reports\ReportTemplate\Report_Template.dotx";
			string newfilePath = @"Reports\GeneratedReports\" + patient.Name + ".docx";
			string vizPath = @"Reports\GeneratedReports\Visualization.docx";
			string imagePath = @"Reports\GeneratedReports\renderedVisualization2.png";

			if (File.Exists(newfilePath))
			{
				File.Delete(newfilePath);
			}

			File.Copy(templatePath, newfilePath);

			
			using(WordprocessingDocument myDoc = WordprocessingDocument.Open(newfilePath,true)){

				myDoc.ChangeDocumentType(DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
				WordAPI.InsertPatientData(myDoc,patient);
				foreach(TestResultGroup testResultGroup in patient.ResultGroups){
					WordAPI.DisplayTestGroup(myDoc,testResultGroup);
				}
				WordAPI.PageBreak(myDoc);
				WordAPI.InsertPicturePng(myDoc, imagePath,7,1.2);
				WordAPI.JoinFile(myDoc,vizPath);
			}

			

			Console.WriteLine("Modified");

		}
	}
}
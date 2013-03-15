using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using org.apache.pdfbox.pdmodel;
using org.apache.pdfbox.util;
using OfficeOpenXml;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace MS2Text
{
    static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            var start = DateTime.Now;
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: <input filepath (PDF,DOC,PPT,XLS)> not required <output filepath (TXT)>");
                return;
            }

            var inputFilePath = args[0];//@"c:\temp\tamer.ppt";
            var textFileName = args[1]; //@"c:\temp\tamerppt.txt";

            if (!string.IsNullOrEmpty(inputFilePath) && File.Exists(inputFilePath))
            {
                Console.WriteLine("Input File Does Not Exist");
                return;
            }

            var fi = new FileInfo(inputFilePath);
            var returnString = string.Empty;
            var result = new List<string>();
            if (fi.Extension == ".pdf")
            {
                returnString = ParseUsingPDFBox(fi.FullName, textFileName);
            }

            if (fi.Extension == ".doc" || fi.Extension == ".docx")
            {
                returnString = ParseWordToText(fi.FullName, textFileName);
            }

            if (fi.Extension == ".xls" || fi.Extension == ".xlsx")
            {
                returnString = ParseExcelToText(fi.FullName, textFileName);
            }

            if (fi.Extension == ".ppt" || fi.Extension == ".pptx")
            {
                returnString = ParsePowerPointToText(fi.FullName, textFileName);
            }

            if (!string.IsNullOrEmpty(returnString))
            {
                result = GetWords(returnString);
            }


            Console.WriteLine(result);
            Console.WriteLine("Done. Took " + (DateTime.Now - start));
            Console.ReadLine();
        }

        private static List<string> GetWords(string inputText)
        {
            return inputText.Split(' ').Where(item => item.Length > 3).ToList();
        }

        private static void KillProcess()
        {
            var GetPArry = Process.GetProcesses();
            foreach (var testProcess in GetPArry)
            {
                var ProcessName = testProcess.ProcessName;

                ProcessName = ProcessName.ToLower();
                if (String.Compare(ProcessName, "powerpnt", StringComparison.Ordinal) == 0)
                    testProcess.Kill();
            }
        }

        private static string ParsePowerPointToText(string inputFile, string outputFile)
        {
            var PowerPoint_App = new PowerPoint.Application();
            try
            {
                var multi_presentations = PowerPoint_App.Presentations;
                var presentation = multi_presentations.Open(inputFile);
                var presentation_text = "";
                for (var i = 0; i < presentation.Slides.Count; i++)
                {
                    presentation_text = (from PowerPoint.Shape shape in presentation.Slides[i + 1].Shapes where shape.HasTextFrame == MsoTriState.msoTrue where shape.TextFrame.HasText == MsoTriState.msoTrue select shape.TextFrame.TextRange into textRange select textRange.Text).Aggregate(presentation_text, (current, text) => current + (text + " "));
                }

                if (string.IsNullOrEmpty(outputFile))
                    return presentation_text;

                using (var sw = new StreamWriter(outputFile))
                {
                    sw.WriteLine(presentation_text);
                }
                Console.WriteLine(presentation_text);
            }
            finally
            {
                PowerPoint_App.Quit();
                Marshal.FinalReleaseComObject(PowerPoint_App);
                GC.Collect();
                KillProcess();
            }
            return string.Empty;
        }

        private static string ParseUsingPDFBox(string inputFile, string outputFile)
        {
            var doc = PDDocument.load(inputFile);
            var stripper = new PDFTextStripper();

            var result = stripper.getText(doc);

            if (string.IsNullOrEmpty(outputFile))
                return result;

            using (var sw = new StreamWriter(outputFile))
            {
                sw.WriteLine(result);
            }

            return string.Empty;
        }

        private static string ParseExcelToText(string inputFile, string outputFile)
        {
            try
            {
                var existingFile = new FileInfo(inputFile);
                using (var package = new ExcelPackage(existingFile))
                {
                    var workBook = package.Workbook;
                    if (workBook == null) return string.Empty;
                    if (workBook.Worksheets.Count > 0)
                    {
                        var currentWorksheet = workBook.Worksheets.First();

                        var result = string.Empty;
                        for (var rowNumber = 1; rowNumber <= currentWorksheet.Dimension.End.Row; rowNumber++)
                        {
                            for (var colNumber = 1;
                                 colNumber <= currentWorksheet.Dimension.End.Column;
                                 colNumber++)
                            {
                                result += currentWorksheet.Cells[rowNumber, colNumber].Value;
                            }
                        }

                        if (string.IsNullOrEmpty(outputFile))
                            return result;

                        using (var sw = new StreamWriter(outputFile))
                        {
                            sw.WriteLine(result);
                        }
                    }
                }
            }

            catch (IOException ioEx)
            {
                if (!String.IsNullOrEmpty(ioEx.Message))
                {
                    if (ioEx.Message.Contains("because it is being used by another process."))
                    {
                        Console.WriteLine("Could not read data. Please make sure it not open in Excel.");
                    }
                }
                Console.WriteLine("Could not read example data. " + ioEx.Message, ioEx);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occured while reading example data. " + ex.Message, ex);
            }
            return string.Empty;
        }

        private static string ParseWordToText(string inputFile, string outputFile)
        {
            try
            {
                var wordApp = new Microsoft.Office.Interop.Word.ApplicationClass();

                string fn = inputFile;

                object oFile = fn;
                object oNull = System.Reflection.Missing.Value;
                object oReadOnly = true;

                var Doc = wordApp.Documents.Open(ref oFile, ref oNull,
                                                           ref oReadOnly, ref oNull, ref oNull, ref oNull, ref oNull,
                                                           ref oNull, ref oNull, ref oNull, ref oNull, ref oNull,
                                                           ref oNull, ref oNull, ref oNull);

                var result = Doc.Paragraphs.Cast<Paragraph>().Aggregate(string.Empty, (current, oPara) => current + oPara.Range.Text);

                using (var sw = new StreamWriter(outputFile))
                {
                    sw.WriteLine(result);
                }

                wordApp.Quit(ref oNull, ref oNull, ref oNull);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return string.Empty;
        }
    }
}

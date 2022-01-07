using iText.IO.Font.Constants;
using iText.Kernel.Font;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using System;
using System.IO;
using System.Collections.Generic;
using iText.Kernel.Pdf.Annot;
using iText.Kernel.Utils;

namespace iText7core
{
    class Program
    {
        static void Main(string[] args)
        {
            #region create pdf from scratch

            if (args[0] == "create")
            {

                using (var writer = new PdfWriter(@"C:\Projects\iText7\output\createoutput.pdf"))
                {
                    using (var pdf = new PdfDocument(writer))
                    {
                        var doc = new Document(pdf);
                        var font = PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN);

                        doc.Add(new Paragraph("This List").SetFont(font));
                        List list = new List()
                            .SetSymbolIndent(12)
                            .SetListSymbol("\u2022")
                            .SetFont(font);

                        list.Add(new ListItem("Page 1"))
                            .Add(new ListItem("Thing 02"))
                            .Add(new ListItem("Thing 03"))
                            .Add(new ListItem("Thing 04"));

                        doc.Add(list);

                        doc.Add(new Paragraph("New Paragraph"));

                        for (int i = 2; i < 100000; i++)
                        {
                            doc.Add(new AreaBreak(AreaBreakType.NEXT_PAGE));

                            var nextPageFont = PdfFontFactory.CreateFont(StandardFonts.COURIER);

                            doc.Add(new Paragraph("New Page").SetFont(nextPageFont));
                            List nextPage = new List()
                                .SetSymbolIndent(12)
                                .SetListSymbol("\u2022")
                                .SetFont(nextPageFont);

                            nextPage.Add(new ListItem($"Page {i}"));
                            doc.Add(nextPage);
                            doc.Add(new Paragraph($"Paragraph on page {i}"));

                        }

                        doc.Close();
                    }
                }
            }

            #endregion

            #region Merge a directory of pdf's

            if (args[0] == "merge")
            {
                string inputDirectory = @"C:\Projects\iText7\input\mergedir\";
                string workDirectory = @"C:\Projects\iText7\workdir\";
                string outputFile = @"C:\Projects\iText7\output\merged.pdf";
                List<string> mergeIndex = new List<string>();

                //code to concat PDF's
                DateTime actionStartTime = DateTime.Now;
                Console.WriteLine($"Starting merge of Pdf's in Directory {inputDirectory}");

                //get the directory listing of pdf's to merge into a list
                DirectoryInfo inputDirectoryInfo = new DirectoryInfo(inputDirectory);
                foreach (FileInfo pdfFile in inputDirectoryInfo.GetFiles("*.pdf"))
                {
                    mergeIndex.Add(pdfFile.FullName);
                }

                using (PdfWriter thisWriter = new PdfWriter(outputFile, new WriterProperties().UseSmartMode()))
                {
                    using (PdfDocument thisDocument = new PdfDocument(thisWriter))
                    {
                        PdfMerger mergedPDF = new PdfMerger(thisDocument);
                        for (int i = 0; i < mergeIndex.Count; i++)
                        {
                            // keep track of what PDF we are running in case it crashes
                            string currentPdf = mergeIndex[i];
                            string annotPdfFile = $@"{workDirectory}{Path.GetFileNameWithoutExtension(mergeIndex[i])}.annot.pdf";

                            try
                            {
                                //add annot to every first page
                                PdfDocument pdfAnnot = new PdfDocument(new PdfReader(currentPdf), new PdfWriter(annotPdfFile));
                                pdfAnnot.GetFirstPage().AddAnnotation(new PdfTextAnnotation(new iText.Kernel.Geom.Rectangle(0, 0, 0, 0))
                                    .SetTitle(new PdfString("BREAK"))
                                    .SetContents(Path.GetFileName(mergeIndex[i])));
                                pdfAnnot.Close();
                                currentPdf = annotPdfFile;

                                //merge the file we added an annot to
                                PdfDocument pdfToMerge = new PdfDocument(new PdfReader(currentPdf));
                                mergedPDF.Merge(pdfToMerge, 1, pdfToMerge.GetNumberOfPages());
                                pdfToMerge.Close();
                            }
                            catch (Exception e)
                            {
                                mergedPDF.Close();
                                Console.WriteLine($"Something went wrong : {e.Message}");
                            }
                        }
                        mergedPDF.Close();
                    }
                }

                Console.WriteLine($"iText PDF Merge completed in {DateTime.Now.Subtract(actionStartTime):c} for Directory {inputDirectory}");
            }

            #endregion
        }
    }
}

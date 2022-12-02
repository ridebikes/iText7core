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
using iText.Kernel.Geom;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Pdf.Xobject;
using iText.Kernel.Pdf.Canvas.Parser;

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
                string inputDirectory = @"C:\Projects\iText7\input\";
                string workDirectory = @"C:\Projects\iText7\workdir\";
                string outputFile = @"C:\Projects\iText7\output\merged.pdf";
                List<string> mergeIndex = new List<string>();

                //code to merge PDF's
                DateTime actionStartTime = DateTime.Now;

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
                            string annotPdfFile = $@"{workDirectory}{System.IO.Path.GetFileNameWithoutExtension(mergeIndex[i])}.annot.pdf";

                            try
                            {
                                //add annot to every first page
                                PdfDocument pdfAnnot = new PdfDocument(new PdfReader(currentPdf), new PdfWriter(annotPdfFile));
                                pdfAnnot.GetFirstPage().AddAnnotation(new PdfTextAnnotation(new iText.Kernel.Geom.Rectangle(0, 0, 0, 0))
                                    .SetTitle(new PdfString("PDF_BREAK"))
                                    .SetContents(System.IO.Path.GetFileName(mergeIndex[i])));
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

                Console.WriteLine($"iText processing completed in {DateTime.Now.Subtract(actionStartTime):c}");
            }

            #endregion

            #region Strip, Flatten and Save a directory of pdf's

            if (args[0] == "stripsave")
            {
                string inputDirectory = @"C:\Projects\iText7\input\";
                string workDirectory = @"C:\Projects\iText7\workdir\";
                string outputDirectory = $@"C:\Projects\iText7\output\";

                List<string> pdfIndex = new List<string>();

                DateTime actionStartTime = DateTime.Now;

                DirectoryInfo inputDirectoryInfo = new DirectoryInfo(inputDirectory);
                foreach (FileInfo pdfFile in inputDirectoryInfo.GetFiles("*.pdf"))
                {
                    pdfIndex.Add(pdfFile.FullName);
                }

                for (int i = 0; i < pdfIndex.Count; i++)
                {
                    // keep track of what PDF we are running in case it crashes
                    string currentPdf = pdfIndex[i];
                    string annotPdfFile = $@"{workDirectory}{System.IO.Path.GetFileNameWithoutExtension(pdfIndex[i])}.annot.pdf";
                    string acroformPdfFile = $@"{workDirectory}{System.IO.Path.GetFileNameWithoutExtension(pdfIndex[i])}.acroform.pdf";

                    //Strip annot from every page
                    PdfDocument pdfAnnotStripper = new PdfDocument(new PdfReader(currentPdf), new PdfWriter(annotPdfFile));

                    for (int j = 1; j <= pdfAnnotStripper.GetNumberOfPages(); j++)
                    {
                        foreach (PdfAnnotation annotation in pdfAnnotStripper.GetPage(j).GetAnnotations())
                        {
                            pdfAnnotStripper.GetPage(j).RemoveAnnotation(annotation);
                        }
                    }
                    pdfAnnotStripper.GetOutlines(true).RemoveOutline();
                    pdfAnnotStripper.Close();
                    currentPdf = annotPdfFile;
                    
                    //flatten Acroforms
                    using (PdfDocument pdfFlattener = new PdfDocument(new PdfReader(currentPdf), new PdfWriter(acroformPdfFile)))
                    {
                        PdfAcroForm thisForm = PdfAcroForm.GetAcroForm(pdfFlattener, true);
                        thisForm.FlattenFields();
                        pdfFlattener.Close();
                        currentPdf = acroformPdfFile;
                    }

                    string outputFile = $@"{outputDirectory}{System.IO.Path.GetFileName(currentPdf)}";

                    // smartmode is a good pdf saver
                    using (PdfWriter thisWriter = new PdfWriter(outputFile, new WriterProperties().UseSmartMode()))
                    {
                        using (PdfDocument thisDocument = new PdfDocument(thisWriter))
                        {
                            PdfMerger mergedPDF = new PdfMerger(thisDocument);
                            PdfDocument sourceDocument = new PdfDocument(new PdfReader(currentPdf));
                            mergedPDF.Merge(sourceDocument, 1, sourceDocument.GetNumberOfPages());
                            sourceDocument.Close();
                            mergedPDF.Close();
                        }
                    }
                }

                Console.WriteLine($"iText processing completed in {DateTime.Now.Subtract(actionStartTime):c}");
            }

            #endregion

            #region ScaleAndRotate

            if (args[0] == "scaleandrotate")
            {
                string inputDirectory = @"C:\Projects\iText7\input\";
                string outputDirectory = $@"C:\Projects\iText7\output\";
                float pageWidth = 612;
                float pageHeight = 792;

                List<string> pdfIndex = new List<string>();
                DateTime actionStartTime = DateTime.Now;

                DirectoryInfo inputDirectoryInfo = new DirectoryInfo(inputDirectory);
                foreach (FileInfo pdfFile in inputDirectoryInfo.GetFiles("*.pdf"))
                {
                    pdfIndex.Add(pdfFile.FullName);
                }

                for (int i = 0; i < pdfIndex.Count; i++)
                {
                    string currentPdf = pdfIndex[i];
                    string outputFile = $@"{outputDirectory}{System.IO.Path.GetFileName(currentPdf)}";

                    // using iText 7.1.12 with Affine Transform
                    // https://api.itextpdf.com/iText7/dotnet/7.1.12/classi_text_1_1_kernel_1_1_geom_1_1_affine_transform.html

                    PageSize targetSize = new PageSize(new Rectangle(pageWidth, pageHeight));

                    using (PdfReader inputReader = new PdfReader(pdfIndex[i]))
                    {
                        using (PdfDocument inputDocument = new PdfDocument(inputReader))
                        {
                            using (PdfWriter outputWriter = new PdfWriter(outputFile))
                            {
                                using (PdfDocument outputDocument = new PdfDocument(outputWriter))
                                {
                                    for (int j = 1; j <= inputDocument.GetNumberOfPages(); j++)
                                    {
                                        PdfPage origPage = inputDocument.GetPage(j);

                                        Rectangle thisRectangle = origPage.GetPageSize();

                                        PageSize thisSize = new PageSize(thisRectangle);

                                        bool needsRotated = (thisSize.GetHeight() >= thisSize.GetWidth()) != (targetSize.GetHeight() >= targetSize.GetWidth());

                                        PageSize rotatedSize = thisSize;

                                        if (needsRotated)
                                        {
                                            rotatedSize = new PageSize(new Rectangle(thisSize.GetHeight(), thisSize.GetWidth()));
                                        }

                                        PdfPage destPage = outputDocument.AddNewPage(targetSize);
                                        AffineTransform transformationMatrix = new AffineTransform();

                                        //double scale = Math.Min(targetSize.GetWidth() / rotatedSize.GetWidth(), targetSize.GetHeight() / rotatedSize.GetHeight());
                                        double scaleX = targetSize.GetWidth() / rotatedSize.GetWidth();
                                        double scaleY = targetSize.GetHeight() / rotatedSize.GetHeight();

                                        //we are going to scale, but to nothing
                                        transformationMatrix = AffineTransform.GetScaleInstance(scaleX, scaleY);

                                        if (needsRotated)
                                        {
                                            //rotate the page (in Radians)
                                            transformationMatrix = AffineTransform.GetRotateInstance(Math.PI / 2, thisSize.GetHeight() / 2, thisSize.GetWidth() / 2);
                                        }

                                        //create canvas
                                        PdfCanvas destCanvas = new PdfCanvas(destPage);

                                        //run Affine config
                                        destCanvas.ConcatMatrix(transformationMatrix);

                                        //grab page data as FormXObject
                                        PdfFormXObject origCopy = origPage.CopyAsFormXObject(outputDocument);

                                        float X = (float)((targetSize.GetWidth() - thisSize.GetWidth() * scaleX) / 2);
                                        float Y = (float)((targetSize.GetHeight() - thisSize.GetHeight() * scaleY) / 2);

                                        //add to our canvas
                                        destCanvas.AddXObjectAt(origCopy, X, Y);

                                        //finalize canvas for next iteration
                                        destCanvas = new PdfCanvas(destPage);
                                    }
                                }
                            }
                        }
                    }

                }

                Console.WriteLine($"iText processing completed in {DateTime.Now.Subtract(actionStartTime):c}");
            }


            #endregion

            #region Convert to text

            if (args[0] == "text")
            {
                string inputDirectory = @"C:\Projects\iText7\input\";
                string outputDirectory = $@"C:\Projects\iText7\output\";

                List<string> pdfIndex = new List<string>();
                DateTime actionStartTime = DateTime.Now;

                DirectoryInfo inputDirectoryInfo = new DirectoryInfo(inputDirectory);
                foreach (FileInfo pdfFile in inputDirectoryInfo.GetFiles("*.pdf"))
                {
                    pdfIndex.Add(pdfFile.FullName);
                }

                for (int i = 0; i < pdfIndex.Count; i++)
                {
                    string currentPdf = pdfIndex[i];
                    string outputFile = $@"{outputDirectory}{System.IO.Path.GetFileName(currentPdf)}";

                    using (PdfDocument inputPDF = new PdfDocument(new PdfReader(currentPdf)))
                    {
                        using (StreamWriter outputText = new StreamWriter(outputFile))
                        {
                            for (int j = 1; j <= inputPDF.GetNumberOfPages(); j++)
                            {
                                var page = inputPDF.GetPage(j);
                                outputText.WriteLine();
                                outputText.WriteLine($"||P{j:0000000000}||");
                                outputText.WriteLine();
                                outputText.WriteLine(PdfTextExtractor.GetTextFromPage(page));
                            }

                            outputText.Close();
                        }
                    }
                }

                Console.WriteLine($"iText processing completed in {DateTime.Now.Subtract(actionStartTime):c}");
            }
            #endregion

            #region Add an Image to a PDF

            if (args[0] == "addimage")
            {
                DateTime actionStartTime = DateTime.Now;

                string inputDirectory = @"C:\Projects\iText7\input\image\";
                string outputDirectory = $@"C:\Projects\iText7\output\image\";
                string imageToAdd = @"C:\Projects\iText7\input\image\myImage.jpg";
                int pageNumber = 1;
                float leftPos = 0;
                float bottomPos = 400;
                float pageHeight;
                float pageWidth;
                float degreesRotation = 0;


                List<string> pdfIndex = new List<string>();
                
                DirectoryInfo inputDirectoryInfo = new DirectoryInfo(inputDirectory);
                foreach (FileInfo pdfFile in inputDirectoryInfo.GetFiles("*.pdf"))
                {
                    pdfIndex.Add(pdfFile.FullName);
                }

                for (int i = 0; i < pdfIndex.Count; i++)
                {
                    string currentPdf = pdfIndex[i];
                    string outputFile = $@"{outputDirectory}{System.IO.Path.GetFileName(currentPdf)}";

                    using (PdfReader inputReader = new PdfReader(pdfIndex[i]))
                    {
                        using (PdfDocument inputDocument = new PdfDocument(inputReader))
                        {
                            using (PdfWriter outputWriter = new PdfWriter(outputFile))
                            {
                                using (PdfDocument outputDocument = new PdfDocument(outputWriter))
                                {
                                    using (PdfReader thisReader = new PdfReader(currentPdf))
                                    {
                                        PdfPage origPage = inputDocument.GetPage(1);
                                        Rectangle thisRectangle = origPage.GetPageSize();
                                        pageHeight = thisRectangle.GetHeight();
                                        pageWidth = thisRectangle.GetWidth();

                                        Document thisDoc = new Document(outputDocument);
                                        Image thisImage = new Image(ImageDataFactory.Create(imageToAdd))
                                            .SetFixedPosition(pageNumber, leftPos, bottomPos)
                                            .SetRotationAngle(degreesRotation * (Math.PI / 180))
                                            .ScaleToFit(pageWidth, pageHeight);
                                        thisDoc.Add(thisImage);
                                        thisDoc.Close();
                                    }
                                }
                            }
                        }
                    }
                    Console.WriteLine($"iText processing completed in {DateTime.Now.Subtract(actionStartTime):c}");
                }
            }


            #endregion    
                
            #region Split a PDF with an index

            if (args[0] == "split")
            {
                string inputFile = @"C:\Projects\iText7\input\filetosplit.pdf";
                List<SplitIndex> splitIndex = new List<SplitIndex>();

                DateTime actionStartTime = DateTime.Now;

                using (PdfDocument inputPdf = new PdfDocument(new PdfReader(inputFile)))
                {

                    foreach (SplitIndex index in splitIndex)
                    {
                        try
                        {
                            var split = new ImprovedSplitter(inputPdf, PageRange => new PdfWriter(index.FileName));
                            var result = split.ExtractPageRange(new PageRange($"{index.FirstPage}-{index.LastPage}"));
                            result.Close();
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine($"Something went wrong : {e.Message}");
                        }
                    }
                }

                Console.WriteLine($"iText processing completed in {DateTime.Now.Subtract(actionStartTime):c}");
            }
        }

        internal class ImprovedSplitter : PdfSplitter
        {
            private Func<PageRange, PdfWriter> nextWriter;
            public ImprovedSplitter(PdfDocument pdfDocument, Func<PageRange, PdfWriter> nextWriter) : base(pdfDocument)
            {
                this.nextWriter = nextWriter;
            }

            protected override PdfWriter GetNextPdfWriter(PageRange documentPageRange)
            {
                return nextWriter.Invoke(documentPageRange);
            }
        }

        #endregion

        internal class SplitIndex
        {
            public int Counter { get; set; }

            public string FileName { get; set; }

            public int FirstPage { get; set; }

            public int LastPage { get; set; }

            public int PageRange
            {
                get
                {
                    return this.LastPage - this.FirstPage;
                }
            }
        }

    }
}

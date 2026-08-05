using System;
using System.Threading;
using System.IO;
using System.Windows.Forms;
using iText.IO.Font.Constants;
using iText.Kernel.Font;
using iText.Kernel.Pdf.Xobject;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Colors;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using iText.Barcodes;
using iText.IO.Image;
using iText.Kernel.Pdf.Annot;

namespace PDFStamper
{
    internal class Program
    {
        static ProgressBarForm progressBarForm;
        static Thread progressBarThread;
        static int Main(string[] args)
        {

            
            //// Start the progress bar form in a new thread
            //StartProgressBar();

            //int totaltest = 100; // Total count for progress bar

            //// Simulate counting work
            //for (int i = 0; i <= totaltest; i++)
            //{
            //    // Simulate some work
            //    Thread.Sleep(50);

            //    // Update the progress bar
            //    UpdateProgressBar(i, totaltest);
            //}

            //// Close the progress bar form
            //CloseProgressBar();

            //return 0;

            #region ArgsCheck
            //Program was run with no / not enough input. Show help menu.
            if ((args.Length == 0 || args.Length < 7) && (args[2].ToString() != "NBP"))
            {
                if (args[2].ToString() != "CONTROL")
                {
                    Intro();
                    return 0;
                }


            }
            else
            {
                if ((args[0] == "/h") || (args[0] == "-h") || (args[0] == "help"))
                {
                    Intro();
                    return 0;
                }
            }
            #endregion

            string Pagetype;
            string DataPacketID;
            string ClientName;
            string ProductName;
            string EventID;
            string InstLocation;
            string InstDesc;
            string SampleNum;
            string InstModel;
            string InstName;
            string[] DocID;
            string Barcode;
            string[] BarcodeList;
            int PDFPageCount;
            string ReportedName;
            bool IsHorizontal;
            int PageNumStart;
            string footertext = "ANALYST INITIALS & DATE:_____________________   REVIEWER INITIALS & DATE:_____________________    PAGE: ";
            try
            {
                //Check page type first as this will determine what other args to look for. If this is missing or invalid nothing else matters.
                //First and Second args should always be input/output file path.
                
                if (string.IsNullOrWhiteSpace(args[2].ToString()))
                {
                    throw new Exception("Page type is invalid!");
                }
                else
                {
                    Pagetype = args[2].ToString();
                }
                string filepath = args[0];
                FileInfo inputfile = new FileInfo(filepath);
                if (!inputfile.Exists)
                {
                    throw new Exception("Input file not found!");
                }
                string pathout = args[1];
                FileInfo outputfile = new FileInfo(pathout);
                outputfile.Delete();

                if (string.IsNullOrWhiteSpace(args[3].ToString()))
                {
                    throw new Exception("Datapacket ID not found!");
                }
                else
                {
                    DataPacketID = args[3].ToString(); //or sample number ... Or Controlled print ID
                }

                if (!int.TryParse(args[4], out PageNumStart))
                {
                    throw new Exception("Page number is invalid!");
                }

                if (string.IsNullOrWhiteSpace(args[5].ToString()))
                {
                    throw new Exception("Reported name not found!");
                }
                else
                {
                    ReportedName = args[5].ToString();
                }

                if (string.IsNullOrWhiteSpace(args[6].ToString()))
                {
                    throw new Exception("Doc ID is invalid");
                }
                else
                {
                    //DocID = args[6].ToString();
                    DocID = args[6].ToString().Split(',');
                }

                PdfDocument pdfDoc;


                if (Pagetype == "PAGECOUNT")
                {
                    PdfReader reader = new PdfReader(inputfile);
                    reader.SetUnethicalReading(true);
                    PdfDocument pdfDoctemp = new PdfDocument(reader);
                    pdfDoc = pdfDoctemp;
                }
                else
                {
                    //Initialize PDF document
                    PdfReader reader = new PdfReader(inputfile);
                    //This has to be set to true for un-passworded docs because PDF's are dumb
                    reader.SetUnethicalReading(true);
                    PdfDocument pdfDoctemp = new PdfDocument(reader, new PdfWriter(outputfile));
                    pdfDoc = pdfDoctemp;
                    
                }

                
                PDFPageCount = pdfDoc.GetNumberOfPages();

                //Valid page types
                //LNBC - Lab Notebook Cover Page
                //ENBC - Equipment Notebook Cover Page
                //ENB - Equipment Notebook Page
                //NBP - Regular notebook page
                //CONTROL - Controlled print barcode only
                //PAGECOUNT - Get the number of pages in a document.
                switch (Pagetype)
                {
                    case "LNBC":
                        {
                            #region ArgsSetup
                            if (string.IsNullOrWhiteSpace(args[7].ToString()))
                            {
                                throw new Exception("Barcode string not found!");
                            }
                            else
                            {
                                Barcode = args[7].ToString();
                            }

                            if (string.IsNullOrWhiteSpace(args[8].ToString()))
                            {
                                throw new Exception("Client name not found!");
                            }
                            else
                            {
                                ClientName = args[8].ToString();
                            }

                            if (string.IsNullOrWhiteSpace(args[9].ToString()))
                            {
                                throw new Exception("Product name not found!");
                            }
                            else
                            {
                                ProductName = args[9].ToString();
                            }
                            #endregion

                            #region LNBC Stamps

                            for (int i = 0; i < PDFPageCount; i++)
                            {


                                PdfPage page = pdfDoc.GetPage(i + 1);
                                PdfCanvas canvasWrite = new PdfCanvas(page);

                                if (ReportedName.ToString() != "BLANK")
                                {
                                    Paragraph p = new Paragraph(ReportedName).SetFontSize(10);
                                    p.SetFont(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN));
                                    p.SetRelativePosition(125, 8, 100, 100);
                                    p.SetMaxWidth(325);
                                    p.SetMaxHeight(40);
                                    new Canvas(page, page.GetPageSize()).Add(p).Close();
                                }


                                
                                canvasWrite.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN), 10).MoveText(475, 770
                                ).ShowText(DataPacketID).EndText();
                                canvasWrite.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN), 10).MoveText(475, 755
                                ).ShowText("Date:_______________").EndText();
                                canvasWrite.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN), 6).MoveText(20, 775
                                ).ShowText(DocID[0]).EndText();
                                PageNumStart++;
                                //vertical footer
                                canvasWrite.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN), 10).MoveText(20, 20
                                   ).ShowText(footertext + (PageNumStart)).EndText();


                                Paragraph title = new Paragraph("Laboratory Notebook Cover Page").SetFontSize(15);
                                title.SetFont(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN));
                                title.SetRelativePosition(190, 40, 100, 100);
                                title.SetMaxWidth(250);
                                new Canvas(page, page.GetPageSize()).Add(title).Close();

                                Barcode128 barcode128 = new Barcode128(pdfDoc);
                                barcode128.SetCodeType(Barcode128.CODE_C);
                                barcode128.SetCode(Barcode);
                                barcode128.SetBarHeight(15);
                                PdfFormXObject barcodeimg = barcode128.CreateFormXObject(ColorConstants.BLACK, ColorConstants.BLACK, pdfDoc);
                                canvasWrite.AddXObjectAt(barcodeimg, 275, 700);

                                canvasWrite.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN), 10).MoveText(20, 695
                                ).ShowText("Client: " + ClientName).EndText();
                                canvasWrite.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN), 10).MoveText(20, 675
                                ).ShowText("Product: " + ProductName).EndText();

                                if (DataPacketID == "QDEMO-DE-MO")
                                {
                                    Paragraph d = new Paragraph("DEMO - Not for production use").SetFontSize(30);
                                    d.SetFont(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN));
                                    //d.SetRelativePosition(220, 150, 100, 100);
                                    d.SetFixedPosition(150, 400, 1000);
                                    //d.SetMaxWidth(525);
                                    d.SetRotationAngle(145);
                                    new Canvas(page, page.GetPageSize()).Add(d).Close();
                                }
                                
                            }
                            #endregion
                            break;

                        }
                    

                    
                    case "ENBC":
                        {
                            #region ArgsSetup

                            EventID = ReportedName;
                            SampleNum = DataPacketID;



                            if (string.IsNullOrWhiteSpace(args[7].ToString()))
                            {
                                throw new Exception("Instrument description not found!");
                            }
                            else
                            {
                                InstDesc = args[7].ToString();
                            }

                            if (string.IsNullOrWhiteSpace(args[8].ToString()))
                            {
                                throw new Exception("Equipment Location not found!");
                            }
                            else
                            {
                                InstLocation = args[8].ToString();
                            }

                            if (string.IsNullOrWhiteSpace(args[9].ToString()))
                            {
                                throw new Exception("Instrument name not found");
                            }
                            else
                            {
                                InstName = args[9].ToString();
                            }

                            if (string.IsNullOrWhiteSpace(args[10].ToString()))
                            {
                                throw new Exception("Instrument model not found");
                            }
                            else
                            {
                                InstModel = args[10].ToString();
                            }
                            #endregion


                            #region ENBC Stamps

                            for (int i = 0; i < PDFPageCount; i++)
                            {
                                PdfPage page = pdfDoc.GetPage(i + 1);
                                PdfCanvas canvasWrite = new PdfCanvas(page);

                                Paragraph p = new Paragraph(InstName).SetFontSize(12);
                                p.SetFont(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN));
                                //p.SetRelativePosition(70, 1, 100, 100);
                                p.SetRelativePosition(75, 5, 100, 100);
                                p.SetMaxWidth(325);
                                new Canvas(page, page.GetPageSize()).Add(p).Close();
                                canvasWrite.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN), 12).MoveText(475, 770
                                ).ShowText(SampleNum).EndText();
                                canvasWrite.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN), 12).MoveText(475, 756
                                ).ShowText(EventID).EndText();
                                canvasWrite.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN), 12).MoveText(475, 742
                                ).ShowText("Date:_______________").EndText();
                                canvasWrite.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN), 6).MoveText(20, 775
                                ).ShowText(DocID[0]).EndText();
                                InstLocation = InstLocation.Substring(0, Math.Min(76,InstLocation.Length));
                                Paragraph et = new Paragraph(InstLocation).SetFontSize(12);
                                et.SetFont(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN));
                                et.SetRelativePosition(75, 47, 100, 100);
                                et.SetMaxWidth(400);
                                new Canvas(page, page.GetPageSize()).Add(et).Close();
                                InstModel = InstModel.Substring(0,Math.Min(76,InstModel.Length));
                                Paragraph model = new Paragraph(InstModel).SetFontSize(12);
                                model.SetFont(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN));
                                model.SetRelativePosition(75, 33, 100, 100);
                                model.SetMaxWidth(400);
                                new Canvas(page, page.GetPageSize()).Add(model).Close();
                                InstDesc = InstDesc.Substring(0, Math.Min(76,InstDesc.Length));
                                Paragraph Cata = new Paragraph(InstDesc).SetFontSize(12);
                                Cata.SetFont(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN));
                                Cata.SetRelativePosition(75, 19, 100, 100);
                                Cata.SetMaxWidth(400);
                                new Canvas(page, page.GetPageSize()).Add(Cata).Close();
                                PageNumStart++;
                                //vertical footer
                                canvasWrite.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN), 10).MoveText(20, 20
                                   ).ShowText(footertext + (PageNumStart)).EndText();

                                if (DataPacketID == "M1111DEMO")
                                {
                                    Paragraph d = new Paragraph("DEMO - Not for production use").SetFontSize(30);
                                    d.SetFont(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN));
                                    //d.SetRelativePosition(220, 150, 100, 100);
                                    d.SetFixedPosition(150, 350, 1000);
                                    //d.SetMaxWidth(525);
                                    d.SetRotationAngle(145);
                                    new Canvas(page, page.GetPageSize()).Add(d).Close();
                                }

                            }
                            #endregion

                            break;
                        }
              

                    case "ENB":
                        {
                            #region ArgsSetup

                            EventID = ReportedName;
                            SampleNum = DataPacketID;



                            if (string.IsNullOrWhiteSpace(args[7].ToString()))
                            {
                                throw new Exception("Instrument description not found!");
                            }
                            else
                            {
                                InstDesc = args[7].ToString();
                            }

                            if (string.IsNullOrWhiteSpace(args[8].ToString()))
                            {
                                throw new Exception("Equipment Location not found!");
                            }
                            else
                            {
                                InstLocation = args[8].ToString();
                            }

                            if (string.IsNullOrWhiteSpace(args[9].ToString()))
                            {
                                throw new Exception("Instrument name not found");
                            }
                            else
                            {
                                InstName = args[9].ToString();
                            }

                            if (string.IsNullOrWhiteSpace(args[10].ToString()))
                            {
                                throw new Exception("Instrument model not found");
                            }
                            else
                            {
                                InstModel = args[10].ToString();
                            }
                            #endregion

                            #region ENB Stamps

                            for (int i = 0; i < PDFPageCount; i++)
                            {
                                PdfPage page = pdfDoc.GetPage(i+1);
                                PdfCanvas canvasWrite = new PdfCanvas(page);

                                Paragraph p = new Paragraph(InstName).SetFontSize(12);
                                p.SetFont(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN));
                                //p.SetRelativePosition(70, 1, 100, 100);
                                p.SetRelativePosition(75, 5, 100, 100);
                                p.SetMaxWidth(325);
                                new Canvas(page, page.GetPageSize()).Add(p).Close();
                                canvasWrite.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN), 12).MoveText(475, 770
                                ).ShowText(SampleNum).EndText();
                                canvasWrite.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN), 12).MoveText(475, 756
                                ).ShowText(EventID).EndText();
                                canvasWrite.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN), 12).MoveText(475, 742
                                ).ShowText("Date:_______________").EndText();
                                canvasWrite.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN), 6).MoveText(20, 775
                                ).ShowText(DocID[i]).EndText();
                                InstLocation = InstLocation.Substring(0, Math.Min(76, InstLocation.Length));
                                Paragraph et = new Paragraph(InstLocation).SetFontSize(12);
                                et.SetFont(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN));
                                et.SetRelativePosition(75, 47, 100, 100);
                                et.SetMaxWidth(400);
                                new Canvas(page, page.GetPageSize()).Add(et).Close();
                                InstModel = InstModel.Substring(0, Math.Min(76, InstModel.Length));
                                Paragraph model = new Paragraph(InstModel).SetFontSize(12);
                                model.SetFont(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN));
                                model.SetRelativePosition(75, 33, 100, 100);
                                model.SetMaxWidth(400);
                                new Canvas(page, page.GetPageSize()).Add(model).Close();
                                InstDesc = InstDesc.Substring(0, Math.Min(76, InstDesc.Length));
                                Paragraph Cata = new Paragraph(InstDesc).SetFontSize(12);
                                Cata.SetFont(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN));
                                Cata.SetRelativePosition(75, 19, 100, 100);
                                Cata.SetMaxWidth(400);
                                new Canvas(page, page.GetPageSize()).Add(Cata).Close();
                                canvasWrite.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN), 12).MoveText(475, 725
                                ).ShowText("Page Not Used").EndText();
                                iText.Kernel.Geom.Rectangle rekt = new iText.Kernel.Geom.Rectangle(565, 722, 14, 14);
                                canvasWrite.Rectangle(rekt);
                                canvasWrite.Stroke();
                                PageNumStart++;
                                //vertical footer
                                canvasWrite.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN), 10).MoveText(20, 20
                                   ).ShowText(footertext + (PageNumStart)).EndText();

                                if (DataPacketID == "M1111DEMO")
                                {
                                    Paragraph d = new Paragraph("DEMO - Not for production use").SetFontSize(30);
                                    d.SetFont(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN));
                                    //d.SetRelativePosition(220, 150, 100, 100);
                                    d.SetFixedPosition(150, 400, 1000);
                                    //d.SetMaxWidth(525);
                                    d.SetRotationAngle(145);
                                    new Canvas(page, page.GetPageSize()).Add(d).Close();
                                }
                            }


                            #endregion

                            break;
                        }


                    case "NBP":
                        {
                            #region NBP Stamps


                            for (int i = 0; i < PDFPageCount; i++)
                            {


                                PdfPage page = pdfDoc.GetPage(i + 1);
                                PdfCanvas canvasWrite = new PdfCanvas(page);
                                iText.Kernel.Geom.Rectangle stuff = page.GetPageSizeWithRotation();
                                if (stuff.GetHeight() >= stuff.GetWidth())
                                {
                                    IsHorizontal = false;
                                }
                                else
                                {
                                    IsHorizontal = true;
                                }
                                

                                PageNumStart++;
                                if (IsHorizontal)
                                {
                                    Paragraph pnu = new Paragraph("Page Not Used").SetFontSize(10);
                                    pnu.SetFont(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN));
                                    pnu.SetFixedPosition(650, 585, 1000);
                                    new Canvas(page, page.GetPageSize()).Add(pnu).Close();
                                    iText.Kernel.Geom.Rectangle rekt = new iText.Kernel.Geom.Rectangle(725, 585, 15, 15);

                                    canvasWrite.Rectangle(rekt);
                                    canvasWrite.Stroke();

                                    if (ReportedName.ToString() != "BLANK")
                                    {
                                        Paragraph p = new Paragraph(ReportedName).SetFontSize(10);
                                        p.SetFont(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN));
                                        //p.SetRelativePosition(770,60, 100, 100);
                                        p.SetFixedPosition(760, 500, 1000);
                                        p.SetRotationAngle(-146.087);
                                        p.SetMaxWidth(325);

                                        new Canvas(page, page.GetPageSize()).Add(p).Close();
                                    }

                                    Paragraph hdocid = new Paragraph(DocID[i]).SetFontSize(6);
                                    hdocid.SetFont(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN));
                                    hdocid.SetFixedPosition(760, 585, 1000);
                                    hdocid.SetRotationAngle(-146.087);
                                    hdocid.SetMaxWidth(300);
                                    new Canvas(page, page.GetPageSize()).Add(hdocid).Close();

                                    Paragraph hdatapacket = new Paragraph(DataPacketID).SetFontSize(10);
                                    hdatapacket.SetFont(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN));
                                    hdatapacket.SetFixedPosition(760, 125, 1000);
                                    hdatapacket.SetRotationAngle(-146.087);
                                    hdatapacket.SetMaxWidth(300);
                                    new Canvas(page, page.GetPageSize()).Add(hdatapacket).Close();

                                    Paragraph hdate = new Paragraph("Date:____________").SetFontSize(10);
                                    hdate.SetFont(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN));
                                    hdate.SetFixedPosition(745, 125, 1000);
                                    hdate.SetRotationAngle(-146.087);
                                    hdatapacket.SetMaxWidth(300);
                                    new Canvas(page, page.GetPageSize()).Add(hdate).Close();

                                    Paragraph hfoot = new Paragraph(footertext + PageNumStart).SetFontSize(10);
                                    hfoot.SetFont(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN));
                                    //hfoot.SetRotationAngle(-146.087);
                                    hfoot.SetRotationAngle(-1.571);
                                    hfoot.SetFixedPosition(6, 600, 1000);
                                    new Canvas(page, page.GetPageSize()).Add(hfoot).Close();
                                }
                                else //vertical stuff
                                {
                                    if (ReportedName.ToString() != "BLANK")
                                    {
                                        Paragraph p = new Paragraph(ReportedName).SetFontSize(10);
                                        p.SetFont(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN));
                                        p.SetRelativePosition(125, 8, 100, 100);
                                        p.SetMaxWidth(325);
                                        p.SetMaxHeight(40);
                                        new Canvas(page, page.GetPageSize()).Add(p).Close();
                                    }

                                    canvasWrite.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN), 10).MoveText(475, 770
                                    ).ShowText(DataPacketID).EndText();
                                    canvasWrite.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN), 10).MoveText(20, 20
                                    ).ShowText(footertext + (PageNumStart)).EndText();
                                    canvasWrite.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN), 10).MoveText(495, 735
                                    ).ShowText("Page Not Used").EndText();
                                    canvasWrite.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN), 10).MoveText(475, 755
                                    ).ShowText("Date:_______________").EndText();
                                    canvasWrite.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN), 6).MoveText(20, 775
                                    ).ShowText(DocID[i]).EndText();

                                    iText.Kernel.Geom.Rectangle rekt = new iText.Kernel.Geom.Rectangle(565, 730, 15, 15);
                                    canvasWrite.Rectangle(rekt);
                                    canvasWrite.Stroke();
                                }

                                if (DataPacketID == "QDEMO-DE-MO")
                                {
                                    Paragraph d = new Paragraph("DEMO - Not for production use").SetFontSize(30);
                                    d.SetFont(PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN));
                                    //d.SetRelativePosition(220, 150, 100, 100);
                                    d.SetFixedPosition(150, 350, 1000);
                                    //d.SetMaxWidth(525);
                                    d.SetRotationAngle(145);
                                    new Canvas(page, page.GetPageSize()).Add(d).Close();
                                }




                            }
                            #endregion

                            break;
                        }

                    case "CONTROL":
                        {
                            #region ArgsSetup

                            if (string.IsNullOrWhiteSpace(args[3].ToString()))
                            {
                                throw new Exception("Control Print ID file string not found");
                            }
                            else
                            {
                                //Barcode = args[3].ToString(); //or Control print ID
                                FileInfo BarcodeIDListFile = new FileInfo(args[3].ToString());
                                if (!BarcodeIDListFile.Exists)
                                {
                                    throw new Exception("Barcode ID file not found!");
                                }
                                BarcodeList = File.ReadAllLines(args[3].ToString());
                                
                            }

                            #endregion

                            #region ControlStamps

                            // Start the progress bar form in a new thread
                            StartProgressBar();

                            int total = PDFPageCount; // Total count for progress bar

                            for (int i = 0; i < PDFPageCount; i++)
                            {
                                UpdateProgressBar(i, total);
                                //pBar.Value = ((i / PDFPageCount) * 100);
                                Barcode = BarcodeList[i];
                                PdfPage page = pdfDoc.GetPage(i + 1);
                                PdfCanvas canvasWrite = new PdfCanvas(page);
                                int barcodelen = Barcode.Length;

                                //Barcode128 barcode128 = new Barcode128(pdfDoc);
                                //barcode128.SetCodeType(Barcode128.CODE_A);
                                //barcode128.SetCode(Barcode);
                                //barcode128.SetSize(0);
                                //barcode128.SetBarHeight(17);
                                //PdfFormXObject barcodeimg = barcode128.CreateFormXObject(ColorConstants.BLACK, ColorConstants.BLACK, pdfDoc);
                                //canvasWrite.AddXObjectAt(barcodeimg, 360, 760);



                                iText.Kernel.Geom.Rectangle stuff = page.GetPageSizeWithRotation();
                                if (stuff.GetHeight() >= stuff.GetWidth())
                                {
                                    IsHorizontal = false;
                                }
                                else
                                {
                                    IsHorizontal = true;
                                    
                                }

                                if (IsHorizontal)
                                {

                                   

                                    Barcode128 barcode128 = new Barcode128(pdfDoc);
                                    barcode128.SetCodeType(Barcode128.CODE_A);
                                    barcode128.SetCode(Barcode);
                                    barcode128.SetSize(0);
                                    barcode128.SetBarHeight(17);
                                    PdfFormXObject barcodeimg = barcode128.CreateFormXObject(ColorConstants.BLACK, ColorConstants.BLACK, pdfDoc);
                                    //PDF coordinates are dumb. reference  "\\qclfilesvr1\LabWare_Admin\Notebook_StampTool_Source_Code\pdfcoords.jpg" and pdfcoords-landscape.jpg for more info.
                                    //also https://www.pdfscripting.com/public/PDF-Page-Coordinates.cfm may help if it is still up.

                                    if (barcodelen > 8)
                                    {//move it way left...
                                        canvasWrite.AddXObjectAt(barcodeimg, 540, 575);
                                    }
                                    else
                                    {//normal position
                                        canvasWrite.AddXObjectAt(barcodeimg, 650, 575);

                                        Paragraph barcodeID = new Paragraph(Barcode).SetFontSize(20);
                                        barcodeID.SetFont(PdfFontFactory.CreateFont(StandardFonts.COURIER_BOLD));
                                        barcodeID.SetFixedPosition(655, 558, 1000);
                                        barcodeID.SetFontColor(ColorConstants.RED);
                                        barcodeID.SetStrokeColor(ColorConstants.RED);
                                        //barcodeID.SetRotationAngle(-1.571);
                                        new Canvas(page, page.GetPageSize()).Add(barcodeID).Close();
                                    }

                                    



                                }
                                else //vertial stuff
                                {
                                    Barcode128 barcode128 = new Barcode128(pdfDoc);
                                    barcode128.SetCodeType(Barcode128.CODE_A);
                                    barcode128.SetCode(Barcode);
                                    barcode128.SetSize(0);
                                    barcode128.SetBarHeight(17);
                                    PdfFormXObject barcodeimg = barcode128.CreateFormXObject(ColorConstants.BLACK, ColorConstants.BLACK, pdfDoc);

                                    if (barcodelen > 8)
                                    {//move it way left...
                                        canvasWrite.AddXObjectAt(barcodeimg, 430, 760);
                                    }
                                    else
                                    {//normal position
                                        canvasWrite.AddXObjectAt(barcodeimg, 485, 760);

                                        Paragraph barcodeID = new Paragraph(Barcode).SetFontSize(17);
                                        barcodeID.SetFont(PdfFontFactory.CreateFont(StandardFonts.COURIER_BOLD));
                                        barcodeID.SetFontColor(ColorConstants.RED);
                                        barcodeID.SetStrokeColor(ColorConstants.RED);
                                        barcodeID.SetFixedPosition(490, 747, 1000);
                                        new Canvas(page, page.GetPageSize()).Add(barcodeID).Close();
                                    }


                                    
                                }








                            }
                            // Close the progress bar form
                            CloseProgressBar();

                            #endregion
                            break;

                        }

                    case "PAGECOUNT":
                        
                            File.WriteAllText(pathout,PDFPageCount.ToString());
                        break;

                    default:
                        throw new Exception("Invalid page type!");
                        
                }
 



                pdfDoc.Close();
                return 1;
            }
            catch (System.IndexOutOfRangeException) 
            {
                Console.WriteLine("Incorrect number of parameters. Notify an Admin ASAP!");
                Intro();
                MessageBox.Show("Incorrect number of parameters found. Notify an Admin ASAP!");
                return 0;

            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
                MessageBox.Show(ex.Message);
                return 0;
            }
           


        }

        // Start the progress bar form in a new thread
        static void StartProgressBar()
        {
            progressBarThread = new Thread(() =>
            {
                progressBarForm = new ProgressBarForm();
                Application.Run(progressBarForm);
                

            });

            progressBarThread.SetApartmentState(ApartmentState.STA);
            progressBarThread.Start();

            // Wait until the form is loaded
            while (progressBarForm == null || !progressBarForm.IsHandleCreated)
            {
                Thread.Sleep(10);
            }
        }

        // Update the progress bar on the form
        static void UpdateProgressBar(int progress, int total)
        {
            progressBarForm.Invoke((Action)(() =>
            {
                progressBarForm.UpdateProgress((progress * 100) / total);
                progressBarForm.Show();
                progressBarForm.BringToFront();
                progressBarForm.Focus();
            }));
        }

        // Close the progress bar form
        static void CloseProgressBar()
        {
            progressBarForm.Invoke((Action)(() =>
            {
                progressBarForm.Close();
            }));

            progressBarThread.Join(); // Wait for the thread to finish
        }

        static void Intro()
        {
            Console.WriteLine("This application takes a PDF and adds a text stamp at pre-defined locations.");
            Console.WriteLine(" ");
            Console.WriteLine("Required Parameters:");
            Console.WriteLine("Input file path");
            Console.WriteLine("Output file path");
            Console.WriteLine("Page Type");
            Console.WriteLine("Data Packet ID");
            Console.WriteLine("Number to start counting pages from");
            Console.WriteLine("Formatted test reported string");
            Console.WriteLine("CSV string of formatted QCL doc ID for notebook pages");
            Console.WriteLine("Cover pages require different info");
            Console.WriteLine(" ");
            Console.WriteLine("Usage example");
            Console.WriteLine(@"PDFStamper.exe C:\temp\demo.pdf c:\temp\output.pdf 'NBP' 'Q0423-1-10' '1' 'USP - Iron' 'NBP - 0043(Ver 1), NBP - 0038(Ver 1)'");
        }

    }

    // Windows Form with a ProgressBar control
    public class ProgressBarForm : Form
    {
        private ProgressBar progressBar;

        public ProgressBarForm()
        {
            this.Text = "Generating Barcodes";
            this.Width = 300;
            this.Height = 100;
            this.CenterToScreen();

            progressBar = new ProgressBar();
            progressBar.Minimum = 0;
            progressBar.Maximum = 100;
            progressBar.Dock = DockStyle.Fill;

            this.Controls.Add(progressBar);
        }

        // Method to update the progress bar value
        public void UpdateProgress(int value)
        {
            progressBar.Value = value;
        }
    }
}

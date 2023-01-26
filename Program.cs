using System;
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

namespace PDFStamper
{
    internal class Program
    {
        static int Main(string[] args)
        {
            #region ArgsCheck
            //Program was run with no / not enough input. Show help menu.
            if (args.Length == 0 || args.Length < 7)
            {
                Intro();
                return 0;
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
            string Equiptype;
            string Location;
            string DocID;
            string Barcode;
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
                    DataPacketID = args[3].ToString();
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
                    DocID = args[6].ToString();
                }

                //Initialize PDF document
                PdfReader reader = new PdfReader(inputfile);
                //This has to be set to true for un-passworded docs because PDF's are dumb
                reader.SetUnethicalReading(true);
                PdfDocument pdfDoc = new PdfDocument(reader, new PdfWriter(outputfile));
                PDFPageCount = pdfDoc.GetNumberOfPages();


                //Valid page types
                //LNBC - Lab Notebook Cover Page
                //ENBC - Equipment Notebook Cover Page
                //ENB - Equipment Notebook Page
                //NBP - Regular notebook page
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
                            PdfCanvas canvasWrite = new PdfCanvas(pdfDoc.GetFirstPage());
                            PdfPage page = pdfDoc.GetFirstPage();

                            Paragraph p = new Paragraph(ReportedName).SetFontSize(10);
                            p.SetFont(PdfFontFactory.CreateFont(StandardFonts.HELVETICA));
                            p.SetRelativePosition(125, 8, 100, 100);
                            p.SetMaxWidth(325);
                            new Canvas(page, page.GetPageSize()).Add(p).Close();
                            canvasWrite.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.HELVETICA), 10).MoveText(475, 770
                            ).ShowText(DataPacketID).EndText();
                            canvasWrite.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.HELVETICA), 10).MoveText(475, 755
                            ).ShowText("Date:_______________").EndText();
                            canvasWrite.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.HELVETICA), 6).MoveText(10, 775
                            ).ShowText(DocID).EndText();

                            //vertical footer
                            canvasWrite.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.HELVETICA), 10).MoveText(10, 15
                               ).ShowText(footertext + (PageNumStart + 1)).EndText();


                            Paragraph title = new Paragraph("Laboratory Notebook Cover Page").SetFontSize(15);
                            title.SetFont(PdfFontFactory.CreateFont(StandardFonts.HELVETICA));
                            title.SetRelativePosition(190, 40, 100, 100);
                            title.SetMaxWidth(250);
                            new Canvas(page, page.GetPageSize()).Add(title).Close();

                            Barcode128 barcode128 = new Barcode128(pdfDoc);
                            barcode128.SetCodeType(Barcode128.CODE_C);
                            barcode128.SetCode(Barcode);
                            barcode128.SetBarHeight(15);
                            PdfFormXObject barcodeimg = barcode128.CreateFormXObject(ColorConstants.BLACK, ColorConstants.BLACK, pdfDoc);
                            canvasWrite.AddXObjectAt(barcodeimg, 275, 700);

                            canvasWrite.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.HELVETICA), 10).MoveText(10, 695
                            ).ShowText("Client: " + ClientName).EndText();
                            canvasWrite.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.HELVETICA), 10).MoveText(10, 675
                            ).ShowText("Product: " + ProductName).EndText();

                            break;
                        }
                    #endregion

                    #region Equipment NB Stamps
                    case "ENBC":
                        {
                            #region ArgsSetup

                            if (string.IsNullOrWhiteSpace(args[7].ToString()))
                            {
                                throw new Exception("Event ID string not found!");
                            }
                            else
                            {
                                EventID = "Event: " + args[7].ToString();
                            }
                            if (string.IsNullOrWhiteSpace(args[8].ToString()))
                            {
                                throw new Exception("Equipment info not found!");
                            }
                            else
                            {
                                Equiptype = args[8].ToString();
                            }

                            if (string.IsNullOrWhiteSpace(args[9].ToString()))
                            {
                                throw new Exception("Location info not found!");
                            }
                            else
                            {
                                Location = "Location: " + args[9].ToString();
                            }
                            #endregion

                            #region ENBC Stamps
                            PdfCanvas canvasWrite = new PdfCanvas(pdfDoc.GetFirstPage());
                            PdfPage page = pdfDoc.GetFirstPage();

                            Paragraph p = new Paragraph("Instrument: " + ReportedName).SetFontSize(8);
                            p.SetFont(PdfFontFactory.CreateFont(StandardFonts.HELVETICA));
                            p.SetRelativePosition(70, 1, 100, 100);
                            p.SetMaxWidth(325);
                            new Canvas(page, page.GetPageSize()).Add(p).Close();
                            canvasWrite.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.HELVETICA), 10).MoveText(475, 780
                            ).ShowText(DataPacketID).EndText();
                            canvasWrite.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.HELVETICA), 10).MoveText(475, 765
                            ).ShowText(EventID).EndText();
                            canvasWrite.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.HELVETICA), 10).MoveText(475, 750
                            ).ShowText("Date:_______________").EndText();
                            canvasWrite.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.HELVETICA), 6).MoveText(10, 775
                            ).ShowText(DocID).EndText();
                            Paragraph et = new Paragraph(Equiptype).SetFontSize(8);
                            et.SetFont(PdfFontFactory.CreateFont(StandardFonts.HELVETICA));
                            et.SetRelativePosition(70, 9, 100, 100);
                            et.SetMaxWidth(325);
                            new Canvas(page, page.GetPageSize()).Add(et).Close();
                            Paragraph Loca = new Paragraph(Location).SetFontSize(8);
                            Loca.SetFont(PdfFontFactory.CreateFont(StandardFonts.HELVETICA));
                            Loca.SetRelativePosition(70, 17, 100, 100);
                            Loca.SetMaxWidth(325);
                            new Canvas(page, page.GetPageSize()).Add(Loca).Close();
                            //vertical footer
                            canvasWrite.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.HELVETICA), 10).MoveText(10, 15
                               ).ShowText(footertext + (PageNumStart + 1)).EndText();

                            #endregion

                            break;
                        }
                    #endregion

                    case "NBP":
                        //So far nothing special goes here?

                        break;
                    default:
                        throw new Exception("Invalid page type!");
                        
                }
 



                pdfDoc.Close();
                return 1;
            }
            catch (System.IndexOutOfRangeException) 
            {
                Console.WriteLine("Incorrect number of parameters. See help menu below.");
                Intro();
                return 0;

            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
                MessageBox.Show(ex.Message);
                return 0;
            }
           


        }

        static void Intro()
        {
            Console.WriteLine("This application takes a PDF and adds a text stamp to it at the given coordinates.");
            Console.WriteLine(" ");
            Console.WriteLine("Required Parameters:");
            Console.WriteLine("Input file path");
            Console.WriteLine("Output file path");
            Console.WriteLine("Text to stamp");
            Console.WriteLine("Page number");
            Console.WriteLine("T/F if page is horizontal");
            Console.WriteLine("Formatted DocID and version");
            Console.WriteLine("Formatted test reported string");
            Console.WriteLine(" ");
            Console.WriteLine("Usage example");
            Console.WriteLine(@"PDFStamper.exe C:\temp\demo.pdf c:\temp\output.pdf 'Text to stamp' 3 F 'NBP-0234 (Ver 1)' 'Evaluation:Dissolution'");
        }
    }
}

using System;
using System.IO;
using System.Windows.Forms;
using iText.IO.Font.Constants;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Colors;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;

namespace PDFStamper
{
    internal class Program
    {
        static int Main(string[] args)
        {
            //Program was run with no / not enough input. Show help menu.
            if (args.Length == 0 || args.Length < 7)
            {
                Intro();
                return 0;
            }
            else
            {
                if ((args[0] == "/h") || (args[0] =="-h") || (args[0] == "help"))
                {
                    Intro();
                    return 0;
                }
            }

            try
            {
                string filepath = args[0];
                FileInfo inputfile = new FileInfo(filepath);
                string pathout = args[1];
                FileInfo outputfile = new FileInfo(pathout);
                outputfile.Delete();


                string DataPacketID;
                string docid;
                bool isHorizontal;
                string ReportedName;

                if (string.IsNullOrWhiteSpace(args[2].ToString()))
                {
                    throw new Exception("Project Number is invalid!");
                }
                else
                {
                    DataPacketID = args[2].ToString();
                }


                if (!int.TryParse(args[3], out _))
                {
                    throw new Exception("Page number is not a number!");
                }

                string footertext = "ANALYST INITIALS & DATE:_____________________   REVIEWER INITIALS & DATE:_____________________    PAGE: " + args[3].ToString();

                if (!inputfile.Exists)
                {
                    throw new Exception("Input file not found!");
                }

                if (args[4].ToString() == "T")
                {
                    isHorizontal = true;
                }
                else if (args[4].ToString() == "F")
                {
                    isHorizontal = false;
                }
                else
                {
                    throw new Exception("Horizontal check failed!");
                }

                if (string.IsNullOrWhiteSpace(args[5].ToString()))
                {
                    throw new Exception("Doc ID and version were invalid!");
                }
                else
                {
                    docid = args[5].ToString();
                }

                if (string.IsNullOrWhiteSpace(args[6].ToString()))
                {
                    throw new Exception("Test reported name was invalid!");
                }
                else
                {
                    ReportedName = args[6].ToString();
                }

                

                //Initialize PDF document
                PdfReader reader = new PdfReader(inputfile);
                //This has to be set to true for un-passworded docs because PDF's are dumb
                reader.SetUnethicalReading(true);
                PdfDocument pdfDoc = new PdfDocument(reader, new PdfWriter(outputfile));
                PdfCanvas canvasWrite = new PdfCanvas(pdfDoc.GetFirstPage());
                PdfPage page = pdfDoc.GetFirstPage();
                Rectangle stuff = page.GetPageSizeWithRotation();
                if (stuff.GetHeight() >= stuff.GetWidth())
                {
                    //page is portrait
                }
                else
                {
                    //page is landscape
                }

                #region Header

                if (isHorizontal)
                {

                }
                else
                {

                    Paragraph p = new Paragraph(ReportedName).SetFontSize(10);
                    p.SetFont(PdfFontFactory.CreateFont(StandardFonts.HELVETICA));
                    p.SetRelativePosition(125, 8, 100, 100);
                    p.SetMaxWidth(325);
                    new Canvas(page, page.GetPageSize()).Add(p).Close();


                    canvasWrite.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.HELVETICA), 10).MoveText(475, 770
                    ).ShowText(DataPacketID).EndText();
                    canvasWrite.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.HELVETICA), 10).MoveText(475, 755
                    ).ShowText("Date:_______________").EndText();
                    canvasWrite.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.HELVETICA), 10).MoveText(495, 739
                    ).ShowText("Page Not Used").EndText();
                    canvasWrite.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.HELVETICA), 6).MoveText(10, 775
                    ).ShowText(docid).EndText();
                    Rectangle rekt = new Rectangle(565, 735, 15, 15);
                    canvasWrite.Rectangle(rekt);
                    canvasWrite.Stroke();


                }

                #endregion

                



                #region Footer
                if (isHorizontal)
                {
                    Paragraph p = new Paragraph(footertext).SetFontSize(10);
                    p.SetFont(PdfFontFactory.CreateFont(StandardFonts.HELVETICA));
                    p.SetRotationAngle(-146.087);
                    p.SetFixedPosition(8, 750, 1000);
                    new Canvas(page, page.GetPageSize()).Add(p).Close();


                }
                else
                {
                    //vertical footer
                    canvasWrite.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.HELVETICA), 10).MoveText(10, 15
                       ).ShowText(footertext).EndText();
                }
                #endregion





                //Use this for horizontal pages
                //Paragraph p = new Paragraph("CONFIDENTIAL").SetFontSize(60);
                //p.SetFont(PdfFontFactory.CreateFont(StandardFonts.HELVETICA));
                //p.SetRotationAngle(90);



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

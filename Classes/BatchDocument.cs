using Batch.Doc;
using CSV_2_PDF;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.Eventing.Reader;
using System.Diagnostics.SymbolStore;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using IO = System.IO;




namespace Batch.Doc
{
    public class ProgressEventArgs : EventArgs
    {
        public int Progress { get; set; }
        public string Message { get; set; }

        public ProgressEventArgs(int progress, string message)
        {
            Progress = progress;
            Message = message;
        }
    }
}

class BatchDocument
    {
        const string NullEntry = "--";
        private Batch.Export.ExportCSV csv = null; 
        public string RepFilename = "Result.pdf";
        public string Title = "Batch Report";
        public string DocumentTitle = "Batch Data";
        public string MachineName = null;
        private int lineCount = 0;
        public int PreTableSize = 0;
        public string Image3PSet = null;
        public string ImageClientSet = null;
        public string FirstColumn = null;
        public string LastColumn = null;

    public int TableStart { get; set; }

    public event EventHandler<ProgressEventArgs> ProgressUpdate;

    public void CreatePdf(Batch.Export.ExportCSV data, string filename)
        {
            csv = data;
            RepFilename = IO.Path.ChangeExtension(filename, "pdf");
            
            Document doc;
            Section section;
            
            doc = new Document();
            
            doc.DefaultPageSetup.PageWidth = "21cm";
            doc.DefaultPageSetup.TopMargin = "3.5cm";
            doc.DefaultPageSetup.LeftMargin = "2cm";
            doc.DefaultPageSetup.RightMargin = "2cm";
            doc.DefaultPageSetup.BottomMargin = "2cm";
  
            DefineStyles(doc);

            section = doc.AddSection();
            Footer(section);
            Header(section, "CellBody8", "CellHeader8");
            Data(section); // includes print table 

            // Render in pdf format
            PdfDocumentRenderer renderer = new PdfDocumentRenderer()
            {
                Document = doc
            };
            renderer.RenderDocument();
            ProgressUpdate(this, e: new ProgressEventArgs(lineCount % 100, "Creating document."));
            // Save the document...
             renderer.PdfDocument.Save(RepFilename);

        }

        /// <summary>
        /// Define the styles used by the document
        /// </summary>
        private void DefineStyles(Document doc)
        {
            // Modify existing styles
            doc.Styles["Normal"].Font.Name = "Verdanna";
            doc.Styles["Normal"].Font.Size = 9;

            doc.Styles["Header"].Font.Bold = true;
            doc.Styles["Header"].ParagraphFormat.Alignment = ParagraphAlignment.Center;
            doc.Styles["Header"].ParagraphFormat.SpaceAfter = "5pt";

            doc.Styles["Footer"].Font.Size = 9;
            doc.Styles["Footer"].Font.Bold = true;
            doc.Styles["Footer"].ParagraphFormat.Alignment = ParagraphAlignment.Center;
            doc.Styles["Footer"].ParagraphFormat.SpaceBefore = "5pt";

            // Create new styles
            Style style;

            style = doc.AddStyle("GridTitle", "Normal");
            style.Font.Size = 12;
            style.Font.Bold = true;
            style.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            style.ParagraphFormat.SpaceAfter = "5pt";

            style = doc.AddStyle("BatchData", "Normal");
            style.Font.Size = 8;
            style.Font.Bold = true;
            style.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            style.ParagraphFormat.SpaceAfter = "2pt";

            style = doc.AddStyle("BatchPreData", "Normal");
            style.Font.Size = 8;
            style.Font.Bold = false;
            style.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            style.ParagraphFormat.SpaceAfter = "2pt";

            style = doc.AddStyle("CellHeader", "Normal");
            style.Font.Size = 10;
            style.Font.Bold = true;
            style.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            style.ParagraphFormat.SpaceAfter = "2pt";
            style.ParagraphFormat.SpaceBefore = "2pt";

            style = doc.AddStyle("CellHeader8", "Normal");
            style.Font.Size = 8;
            style.Font.Bold = true;
            style.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            style.ParagraphFormat.SpaceAfter = "2pt";
            style.ParagraphFormat.SpaceBefore = "2pt";

            style = doc.AddStyle("CellBody", "Normal");
            style.Font.Size = 10;
            style.Font.Bold = true;
            style.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            style.ParagraphFormat.LeftIndent = "0.5cm";
            style.ParagraphFormat.SpaceAfter = "2pt";
            style.ParagraphFormat.SpaceBefore = "2pt";

            style = doc.AddStyle("CellBody8", "Normal");
            style.Font.Size = 8;
            style.Font.Bold = false;
            style.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            style.ParagraphFormat.SpaceAfter = "1pt";
            style.ParagraphFormat.SpaceBefore = "1pt";

            style = doc.AddStyle("CellBody8C", "Normal");
            style.Font.Size = 8;
            style.Font.Bold = false;
            style.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            style.ParagraphFormat.SpaceAfter = "1pt";
            style.ParagraphFormat.SpaceBefore = "1pt";

            style = doc.AddStyle("Title", "Normal");
            style.Font.Size = 20;
            style.Font.Bold = true;
            style.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            style.ParagraphFormat.SpaceAfter = "1pt";
            style.ParagraphFormat.SpaceBefore = "1pt";

        }

        /// <summary>
        /// Create the document header
        /// </summary>
        /// <param name="section">The section to apply the header to.</param>
        private void Header(Section section, String Style, String StyleHead)
        {
            SectionTitle(section, DocumentTitle);
            MigraDoc.DocumentObjectModel.Shapes.Image ImageClient;
            MigraDoc.DocumentObjectModel.Shapes.Image Image3P;
            MigraDoc.DocumentObjectModel.Tables.Table ImageTable;
            MigraDoc.DocumentObjectModel.Tables.Row row;
            Paragraph paragraph;


            ImageTable = section.Headers.Primary.AddTable();
            ImageTable.AddColumn("15cm");
            ImageTable.AddColumn("2cm");

            string ClientLogoPath = ImageClientSet;

            //IO.Path.Combine(IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "3Plogo.bmp");
            //..../calling assembly = 
            row = ImageTable.AddRow();
            paragraph = row.Cells[0].AddParagraph();
            paragraph.AddImage(ClientLogoPath);
            paragraph.Format.SpaceBefore = "0.3cm";

            //  ImageClient.LockAspectRatio = false;
            // ImageClient.Height = "0.8cm";
            // ImageClient.Width = "6.263cm";
            // ImageClient.WrapFormat.Style = MigraDoc.DocumentObjectModel.Shapes.WrapStyle.Through;


            string LogoPath3P = Image3PSet;
            paragraph = row.Cells[1].AddParagraph();
            paragraph.AddImage(LogoPath3P);

            //IO.Path.Combine(IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "3Plogo.bmp");
            //..../calling assembly = 
         /*   Image3P = section.Headers.Primary.AddImage(LogoPath3P);
            Image3P.LockAspectRatio = false;
            Image3P.Height = "2cm";
            Image3P.Width = "3cm";
            Image3P.WrapFormat.Style = MigraDoc.DocumentObjectModel.Shapes.WrapStyle.Through;
         */
        /*
            paragraph = section.Headers.Primary.AddParagraph();
           
     

            paragraph.Format.SpaceBefore = "0.5cm";
            paragraph.Format.Alignment = ParagraphAlignment.Right;

            paragraph = section.Headers.Primary.AddParagraph();
            paragraph.AddText(Title);
            paragraph.Format.Alignment = ParagraphAlignment.Center;
        */
    }

    /// <summary>
    /// Create the document footer
    /// </summary>
    /// <param name="section"></param>
    private void Footer(Section section)
        {
            MigraDoc.DocumentObjectModel.Tables.Table table;
            MigraDoc.DocumentObjectModel.Tables.Row row;
            Paragraph paragraph;

            table = section.Footers.Primary.AddTable();
            table.AddColumn("17.2cm");

            row = table.AddRow();
            row.Borders.Top.Style = BorderStyle.Single;
            row.Borders.Top.Width = "2pt";
            row.Borders.Top.Color = Colors.Blue;

            paragraph = row.Cells[0].AddParagraph();
            paragraph.AddText(String.Format("Printed on {0:dd/MMM/yyyy HH:mm}    Ref:{1}", DateTime.Now, RepFilename));
            paragraph.Style = "Footer";
        }

        /// <summary>
        /// Create Data Section 
        /// </summary>
        /// <param name="section"></param>
    private void Data(Section section)
        {
            MigraDoc.DocumentObjectModel.Tables.Table pretable;
            MigraDoc.DocumentObjectModel.Tables.Row row;
            Paragraph paragraph;


            pretable = section.AddTable();
            pretable.AddColumn("2.3cm");
            pretable.AddColumn("7cm");
            pretable.KeepTogether = false;

            row = pretable.AddRow();
            row.Style = "CellHeader8";
            paragraph = row.Cells[0].AddParagraph();
            paragraph.AddText(String.Format("Machine"));
            paragraph.Style = "BatchData";
            paragraph = row.Cells[1].AddParagraph();
            paragraph.AddText(String.Format(MachineName));
            paragraph.Style = "BatchPreData";

        //  row.Borders.Top.Style = BorderStyle.Single;
        // row.Borders.Top.Width = "2pt";
        // row.Borders.Top.Color = Colors.Blue;

        foreach (var (item, index) in csv.csvList.Select((item, index) => (item, index)))
            {
                if (item == "Batch Code")
                    {
                        row = pretable.AddRow();
                        paragraph = row.Cells[0].AddParagraph();
                        paragraph.AddText(String.Format(item));
                        paragraph.Style = "BatchData";

                        var IndexData = csv.csvList[index+1];

                        paragraph = row.Cells[1].AddParagraph();
                        paragraph.AddText(String.Format(IndexData));
                        paragraph.Style = "BatchPreData";
                    }
                else if (item == "Opened by")
                    {
                        row = pretable.AddRow();
                        paragraph = row.Cells[0].AddParagraph();
                        paragraph.AddText(String.Format(item));
                        paragraph.Style = "BatchData";

                        var IndexData = csv.csvList[index + 1];

                        paragraph = row.Cells[1].AddParagraph();
                        paragraph.AddText(String.Format(IndexData));
                        paragraph.Style = "BatchPreData";
                    }
                else if (item == "Opened at")
                {
                    row = pretable.AddRow();
                    paragraph = row.Cells[0].AddParagraph();
                    paragraph.AddText(String.Format(item));
                    paragraph.Style = "BatchData";

                    var IndexData = csv.csvList[index + 1];

                    paragraph = row.Cells[1].AddParagraph();
                    paragraph.AddText(String.Format(IndexData));
                    paragraph.Style = "BatchPreData";
                }
                else if (item == "Closed by")
                {
                    row = pretable.AddRow();
                    paragraph = row.Cells[0].AddParagraph();
                    paragraph.AddText(String.Format(item));
                    paragraph.Style = "BatchData";

                    var IndexData = csv.csvList[index + 1];

                    paragraph = row.Cells[1].AddParagraph();
                    paragraph.AddText(String.Format(IndexData));
                    paragraph.Style = "BatchPreData";
                }
                else if (item == "Closed at") //May be redundant by the first At
                {
                    row = pretable.AddRow();
                    paragraph = row.Cells[0].AddParagraph();
                    paragraph.AddText(String.Format(item));
                    paragraph.Style = "BatchData";

                    var IndexData = csv.csvList[index + 1];

                    paragraph = row.Cells[1].AddParagraph();
                    paragraph.AddText(String.Format(IndexData));
                    paragraph.Style = "BatchPreData";
                }
                else if (item == "Plate Code") //May be redundant by the first At
                {
                    row = pretable.AddRow();
                    paragraph = row.Cells[0].AddParagraph();
                    paragraph.AddText(String.Format(item));
                    paragraph.Style = "BatchData";

                    var IndexData = csv.csvList[index + 1];

                    paragraph = row.Cells[1].AddParagraph();
                    paragraph.AddText(String.Format(IndexData));
                    paragraph.Style = "BatchPreData";
                }
        }


        PrintDataTable(section, "CellBody8", "CellHeader8");

    }

        /// <summary>
        /// Format chronological events table
        /// </summary>
        /// <param name="section"></param>
        /// <param name="Style"></param>
        /// <param name="StyleHead"></param>
        private void PrintDataTable(Section section, String Style, String StyleHead)
        {
            MigraDoc.DocumentObjectModel.Tables.Table table;
            MigraDoc.DocumentObjectModel.Tables.Table pretable;
            MigraDoc.DocumentObjectModel.Tables.Row row;
            Paragraph paragraph;



            table = section.AddTable();
            table.Borders.Color = Colors.Blue;
            table.Borders.Visible = false;
            table.KeepTogether = false;
        // Probably needs to be in the for loop now that we need configurable column widths 
            // Create columns
            table.AddColumn("1.5cm"); //Dose Position
            table.AddColumn("2.5cm"); //Container Code 
            table.AddColumn("2.5cm"); //Powder Code 
            table.AddColumn("1.8cm"); //Pre Weight
            table.AddColumn("1.8cm"); //Post Weight
            table.AddColumn("1.8cm"); //Delta Weight
            table.AddColumn("1.8cm"); //Pre Weight2 
            table.AddColumn("1.8cm"); //Post Weight2 
            table.AddColumn("1.8cm"); //Delta Weight2 
    
            // Create header row
            row = table.AddRow();
            row.Style = StyleHead;
            row.HeadingFormat = true;
            row.Borders.Top.Width = 1;
            row.Borders.Top.Visible = true;
            row.Borders.Bottom.Width = 2;
            row.Borders.Bottom.Visible = true;
            row.Cells[0].AddParagraph("Plate Position");
            row.Cells[1].AddParagraph("Container Code");
            row.Cells[2].AddParagraph("Sample Code");
            row.Cells[3].AddParagraph("Pre-Dose Weight");
            row.Cells[4].AddParagraph("Post-Dose Weight");
            row.Cells[5].AddParagraph("Delta  Weight");
            row.Cells[6].AddParagraph("Pre-Dose Weight 2");
            row.Cells[7].AddParagraph("Post-Dose Weight 2");
            row.Cells[8].AddParagraph("Delta Weight 2");
        // Create rows

        int counter = 0;
        int columnnumber = 0;
        int HeaderStart = 0;
        int HeaderEnd = 0;
        foreach (var (item, index) in csv.csvList.Select((item, index) => (item, index)))
        {
            if (item == FirstColumn)
                    {
                HeaderStart = index; 
            }
            else if (item == LastColumn)
            {
                HeaderEnd = index;
                columnnumber = HeaderEnd - HeaderStart + 1; // add the 1 cos why not, extra is always bonus 
                counter++;
            }
            else if (counter > 0)
            {
                if (counter % columnnumber == 1) //create row first
                {
                    row = table.AddRow();
                    row.Style = Style;
                    row.Cells[0].AddParagraph(item);
                    counter++;
                }
                else if (counter % columnnumber == 0) // exception for last column as mod = 0
                {
                    row.Style = Style;
                    row.Cells[(columnnumber - 1)].AddParagraph(item);
                    lineCount++;
                    counter++;
                }
                else
                {
                    for (int i = 2; (i < columnnumber); i++) // populate middle rows through loop
                    {
                        if (counter % columnnumber == i)
                        {
                            row.Style = Style;
                            row.Cells[(i-1)].AddParagraph(item);
                        }
                    }
                    counter++;
                }

                row.Borders.Bottom.Width = 1;
                row.Borders.Bottom.Visible = true;
                // Insert zero width space after ever. in the variable name to split the line for long tags
            }
        }
        
            // Vertical line formatting
            foreach (MigraDoc.DocumentObjectModel.Tables.Column col in table.Columns)
            {
                col.Borders.Right.Width = 1;
                col.Borders.Right.Visible = true;
            }
            table.Columns[0].Borders.Left.Width = 1;
            table.Columns[0].Borders.Left.Visible = true;
        
        }


        /// <summary>
        /// Add title paragraph to section
        /// </summary>
        /// <param name="section"></param>
        /// <param name="title"></param>
        private void SectionTitle(Section section, String title)
        {
            Paragraph paragraph;

            paragraph = section.AddParagraph();
            paragraph.AddText(title);
            paragraph.Style = "Title";
            paragraph.Format.SpaceBefore = "0cm";
            paragraph.Format.SpaceAfter = "0.5cm";
        }
}

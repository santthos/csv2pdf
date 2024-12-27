using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CSV_2_PDF
{
    /// <summary>
    /// 3P AML Report generator
    /// </summary>
    class Program
    {
        static frmProgress frm = new frmProgress();
        /// <summary>
        /// Entry point for program
        /// </summary>
        /// <param name="args">Command arguements. See Config for details</param>
        static void Main(string[] args)
        {
            // Initialise configuration
            Config config = new Config
            {
                // Default location for files
                IniFileName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "csv2pdf.ini")
              /*  csvFileName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "BatchDataTest.csv"),
                ReportPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)),
                PreTableSize = 63 */
            }; 
            // Read ini file for remaining configuration
            _ = config.ReadIniFile();

            frm.StartPosition = FormStartPosition.CenterScreen;
            frm.TopMost = true;
            frm.Show();
            try
            {
                // Read in exported Data file
                string csvFilename = config.csvFileName;
                Batch.Export.ExportCSV csv = new Batch.Export.ExportCSV();

                csv.DateTimeFormat = "yyyyMMddTHHmmss.ff";
                frm.Message = "Reading file";
                
                if (csv.Read(csvFilename))
                {
                    frm.Progress = 5;
                    File.Move(csvFilename, Path.Combine(config.ReportPath, String.Format("BatchDataTest.csv")));//("BatchLog{0:yyyMMddHHmmss}.xml", DateTime.Now)));
                    string ReportFileName = Path.Combine(config.ReportPath, String.Format("BatchReport{0:yyyMMddHHmmss}", DateTime.Now));
                    // Create document // transferring the ini file stuff to the document
                    BatchDocument document = new BatchDocument();
                    document.PreTableSize = config.PreTableSize; 
                    document.DocumentTitle = config.DocumentTitle;
                    document.Image3PSet = config.Logo3P;
                    document.ImageClientSet = config.logoClient;
                    document.MachineName = config.MachineName;
                    document.FirstColumn = config.FirstColumnHeader;
                    document.LastColumn = config.LastColumnHeader;

                    document.ProgressUpdate += Document_ProgressUpdate;
                    frm.Message = "Creating report";
                    frm.Progress = 10;

                    document.Title = Path.GetFileNameWithoutExtension(csvFilename).ToUpper();
                   // document.DocumentTitle = Properties.Resources.DocumentTitle;

                    document.CreatePdf(csv, ReportFileName);

                   // Display document
                    //if (config.DisplayDoc && File.Exists(document.RepFilename))
                    Process.Start(document.RepFilename);
                    frm.Progress = 100;
                    frm.Message = "Document completed.";
                }
                else
                {
                    frm.Message = "File not found";
                }
            }
            catch (Exception) { }
            Thread.Sleep(1200);
            frm.Close();
        }


        private static void Document_ProgressUpdate(object sender, Batch.Doc.ProgressEventArgs e)
        {
            frm.Message = e.Message;
            frm.Progress = e.Progress;
        }

    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSV_2_PDF
{
    class Config
    {

        public String ReportPath { get; set; }
        public String IniFileName { get; set; }

        public String csvFileName { get; set; }
        public String Logo3P { get; set; }
        public String logoClient { get; set; }

        public String DocumentTitle { get; set; } 
        public String MachineName { get; set; }
        public int PreTableSize { get; set; } // Number of cells before data table begins

        public string FirstColumnHeader { get; set; } // name of first header
        public string LastColumnHeader { get; set; } // name of last header used to calculate column number

        public Boolean DisplayDoc { get; set; } // Display reort when created, for debug

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]

        public Boolean ReadIniFile()
        {
            // Test if file exists
            if (!File.Exists(IniFileName))
            {
                return false;
            }

            // Open the file to read from
            StreamReader reader = new StreamReader(IniFileName);
            String LineString;
            String[] splitstrings;
            String section = "";

            while (!reader.EndOfStream)
            {
                LineString = reader.ReadLine();

                // Trim spaces and discard comments (after ;)
                LineString = LineString.Trim().Split(';')[0].Trim();

                // Ignore blank lines
                if (!string.IsNullOrEmpty(LineString))
                {
                    splitstrings = LineString.Split('=');
                    if (splitstrings[0].StartsWith("["))
                        section = splitstrings[0].TrimStart('[').TrimEnd(']');
                    else
                    {
                        if (splitstrings.GetUpperBound(0) > 0)
                        {
                            DecodeEntry(section, splitstrings[0].Trim(), splitstrings[1].Trim());
                        }
                    }
                }
            }
            reader.Close();

            return true;
        }

        /// <summary>
        /// Decode line of inifile
        /// </summary>
        /// <param name="section">Current section of ini file</param>
        /// <param name="item">Item name</param>
        /// <param name="value">Item value</param>
        private void DecodeEntry(String section, String item, String value)
        {
            switch (section)
            {
                case "":
                    switch (item.ToUpper()) // will return uppercase
                    {
                        case "":

                            break;

                        case "REPORTPATH":
                            if (!string.IsNullOrEmpty(value))
                                ReportPath = value;
                            break;

                        case "CSVPATH":
                            if (!string.IsNullOrEmpty(value))
                                csvFileName = value;
                            break;

                        case "PREDATATABLESIZE":
                            if (!string.IsNullOrEmpty(value))
                                PreTableSize = Convert.ToInt32(value); //Convert string to int
                            break;

                        case "LOGO3P":
                            if (!string.IsNullOrEmpty(value))
                                Logo3P = value;
                            break;

                        case "LOGOCLIENT":
                            if (!string.IsNullOrEmpty(value))
                                logoClient = value;
                            break;

                        case "DISPLAYDOC":
                            if (!string.IsNullOrEmpty(value))
                                DisplayDoc = Convert.ToBoolean(value);
                            break;

                        case "TITLE":
                            if (!string.IsNullOrEmpty(value))
                                DocumentTitle = value;
                            break;

                        case "MACHINENAME":
                            if (!string.IsNullOrEmpty(value))
                                MachineName = value;
                            break;

                        case "FIRSTCOLUMNHEADER":
                            if (!string.IsNullOrEmpty(value))
                                FirstColumnHeader = value;
                            break;

                        case "LASTCOLUMNHEADER":
                            if (!string.IsNullOrEmpty(value))
                                LastColumnHeader = value;
                            break;
                    }

                    break;
            }






        }
    }
}




using MigraDoc.DocumentObjectModel.Tables;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using CsvHelper;
using CsvHelper.Configuration;
using CsvHelper.Configuration.Attributes;
using System.Windows.Forms;
using MigraDoc.DocumentObjectModel;
using System.Security.Cryptography.X509Certificates;

namespace Batch.Export
{
    class ExportCSV
    {
        public string DateTimeFormat = "dd-MMM-yy HH:mm:ss";

        public List<string> csvList = new List<string>();
        public bool Read(string filename)
        {
            if (string.IsNullOrWhiteSpace(filename))
                return false;
            if (!File.Exists(filename))
                return false;

            string filepath = filename;
            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                HasHeaderRecord = true,
            };
            StreamReader reader = new StreamReader(filepath);

            using (var csv = new CsvReader(reader, config))
            {
                while (csv.Read())
                {
                    int MaxIndex = csv.Parser.Count;
                    for (var i = 0; i < MaxIndex; i++)
                    {
                        csvList.Add(csv.GetField<string>(i));
                    }

                }      
            }
            return true;
        }

    }
}

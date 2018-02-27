using CsvHelper;
using CsvHelper.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ShopifyCSVConverter
{
    public partial class Converter
    {
        //method courtesy of Josip Kremenic https://www.codeproject.com/Articles/231582/Auto-detect-CSV-separator
        public static char Detect(TextReader reader, int rowCount, IList<char> separators)
        {
            IList<int> separatorsCount = new int[separators.Count];
            int character;
            int row = 0;
            bool quoted = false;
            bool firstChar = true;

            while (row < rowCount)
            {
                character = reader.Read();
                switch (character)
                {
                    case '"':
                        if (quoted)
                        {
                            if (reader.Peek() != '"') // Value is quoted and 
                                                      // current character is " and next character is not ".
                                quoted = false;
                            else
                                reader.Read(); // Value is quoted and current and 
                                               // next characters are "" - read (skip) peeked qoute.
                        }
                        else
                        {
                            if (firstChar)  // Set value as quoted only if this quote is the 
                                            // first char in the value.
                                quoted = true;
                        }
                        break;
                    case '\n':
                        if (!quoted)
                        {
                            ++row;
                            firstChar = true;
                            continue;
                        }
                        break;
                    case -1:
                        row = rowCount;
                        break;
                    default:
                        if (!quoted)
                        {
                            int index = separators.IndexOf((char)character);
                            if (index != -1)
                            {
                                ++separatorsCount[index];
                                firstChar = true;
                                continue;
                            }
                        }
                        break;
                }

                if (firstChar)
                    firstChar = false;
            }

            int maxCount = separatorsCount.Max();
            return maxCount == 0 ? '\0' : separators[separatorsCount.IndexOf(maxCount)];
        }

        public DataTable LoadCsv()
        {
            char delimiter;
            using (var textReader = File.OpenText(OpenCsvPath))
            {
                delimiter = Detect(textReader, 5, new char[] { char.Parse(","), char.Parse("|"), char.Parse(":"), char.Parse(";"), char.Parse("\t") });                
            }
            using (var reader = File.OpenText(OpenCsvPath))
            using (var csv = new CsvParser(reader))
            {                                    
                csv.Configuration.Delimiter = delimiter.ToString();
                csv.Configuration.BadDataFound = null;
                csv.Configuration.BufferSize = 4096;
                string[] headers = csv.Read();
                if (headers != null && headers.Length > 0)
                {
                    try
                    {
                        var table = new DataTable();
                        for (int i = 0; i < headers.Length; i++)
                        {
                            var header = headers[i];
                            table.Columns.Add(header);
                        }
                        table.BeginLoadData();
                        while (true)
                        {
                            var record = csv.Read();
                            if (record == null || record.Length == 0) break;                                
                            table.LoadDataRow(record, true);
                        }
                        table.EndLoadData();
                        return table;
                    }
                    catch (Exception) { }
                }
                return null;                
            }
        }
        //Method courtesy of Garran https://stackoverflow.com/questions/31974538/converting-numbers-to-excel-letter-column-vb-net#_=_ (I suck at maths(and coding))
        public static string GetColumnName(int index)
        {
            if (index < 0 || index > 100) return string.Empty;
            var dividend = index + 1;
            string columnName = string.Empty;
            int modulus;
            while (dividend > 0)
            {
                modulus = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulus).ToString() + columnName;
                dividend = (dividend - modulus) / 26;
            }
            return columnName;
        }

        private async void SaveCsv1()
        {
            toolStripStatusLabel1.Text = "Saving";
            toolStripProgressBar1.Style = ProgressBarStyle.Marquee;
            toolStripProgressBar1.Visible = true;
            await Task.Run(() => 
            {

                using (var dataTable = new DataTable())
                using (var reader = new DataTableReader(newDataTable))
                using (var stream = File.Create(SaveCsvPath, 4096, FileOptions.SequentialScan))
                using (var csv = new CsvWriter(new StreamWriter(stream), new Configuration()
                {
                    Encoding = Encoding.UTF8,
                    BufferSize = 4096,
                    TrimOptions = TrimOptions.Trim | TrimOptions.InsideQuotes
                }))
                {
                    try
                    {
                        dataTable.Load(reader);
                        foreach (DataColumn column in dataTable.Columns)
                        {
                            csv.WriteField(column.ColumnName);
                        }
                        csv.NextRecord();
                        foreach (DataRow row in dataTable.Rows)
                        {
                            for (var i = 0; i < dataTable.Columns.Count; i++)
                            {
                                csv.WriteField(row[i]);
                            }
                            csv.NextRecord();
                        }
                        csvNeedsSave = false;
                    }
                    catch (Exception) { }
                }
            });
            toolStripProgressBar1.Visible = false;
            toolStripStatusLabel1.Text = "Ready";
        }

        private async void SaveCsv()
        {
            toolStripProgressBar1.Style = ProgressBarStyle.Marquee;
            toolStripStatusLabel1.Text = "Saving";
            toolStripProgressBar1.Visible = true;
            await Task.Run(() => 
            {
                using (var dataTable = new DataTable())
                using (var reader = new DataTableReader(newDataTable))
                using (var stream = File.Create(SaveCsvPath))
                using (var csv = new CsvWriter(new StreamWriter(stream), new Configuration()
                {
                    Encoding = Encoding.UTF8,
                    BufferSize = 4096,
                    TrimOptions = TrimOptions.Trim | TrimOptions.InsideQuotes
                }))
                {
                    csv.Configuration.RegisterClassMap<ShopifyCsvRowMap>();
                    dataTable.Load(reader);
                    var headerWritten = false;
                    foreach (DataRow row in dataTable.Rows)
                    {
                        try
                        {

                            var shopifyRow = new ShopifyCsvRow(row);
                            if (!headerWritten)
                            {
                                csv.WriteHeader<ShopifyCsvRow>();
                                headerWritten = true;
                            }
                            else csv.WriteRecord(shopifyRow);
                            shopifyRow.Dispose();
                            csv.NextRecord();
                        }
                        catch (Exception) { }
                    }
                }
            });
            csvNeedsSave = false;
            toolStripStatusLabel1.Text = "Ready";
            toolStripProgressBar1.Visible = false;
        }
    }
}

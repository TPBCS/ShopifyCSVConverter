using CsvHelper;
using CsvHelper.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
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

        public async Task<DataTable> LoadCsv()
        {
            char delimiter;
            using (var textReader = File.OpenText(OpenCsvPath))
            {
                delimiter = Detect(textReader, 5, ",|\t#.:;".ToCharArray());                
            }
            using (var reader = File.OpenText(OpenCsvPath))
            using (var csv = new CsvParser(reader))
            {                                    
                csv.Configuration.Delimiter = delimiter.ToString();
                csv.Configuration.BadDataFound = null;
                csv.Configuration.BufferSize = 4096;
                string[] headers = await csv.ReadAsync();
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
                        #if DEBUG
                        var stopWatch = new Stopwatch();
                        stopWatch.Start();
                        #endif
                        while (true)
                        {
                            var record = await csv.ReadAsync();
                            if (record == null || record.Length == 0) break;                                
                            table.LoadDataRow(record, true);
                        }
                        #if DEBUG
                        stopWatch.Stop();
                        MessageBox.Show($"It took: {stopWatch.Elapsed.ToString()}");
                        #endif
                        table.EndLoadData();
                        return table;
                    }
                    catch (Exception) { }
                }
                return null;                
            }
        }

        public async Task<DataTable> LoadCsv1()
        {
            char delimiter;
            using (var textReader = File.OpenText(OpenCsvPath))
            {
                delimiter = Detect(textReader, 5, ",|\t#.:;".ToCharArray());
            }
            using (var reader = File.OpenText(OpenCsvPath))
            using (var csv = new CsvReader(reader, new Configuration() { Delimiter = delimiter.ToString(), BadDataFound = null, BufferSize = 4096}))
            {
                try
                {
                    var dataTable = new DataTable();
                    csv.Read();
                    csv.ReadHeader();
                    foreach (var header in csv.Context.HeaderRecord)
                    {
                        dataTable.Columns.Add(header);
                    }
                    dataTable.BeginLoadData();
                    #if DEBUG
                    var stopWatch = new Stopwatch();
                    stopWatch.Start();
                    #endif
                    while (await csv.ReadAsync())
                    {
                        var row = new string[dataTable.Columns.Count];
                        for (int i = 0; i < row.Length; i++)
                        {
                            row[i] = csv.GetField<string>(i);
                        }
                        dataTable.LoadDataRow(row, true);
                    };
                    #if DEBUG
                    stopWatch.Stop();
                    MessageBox.Show($"It took: {stopWatch.Elapsed.ToString()}");
                    #endif
                    dataTable.EndLoadData();
                    return dataTable;
                }
                catch (Exception) { }
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

        private async Task<bool> SaveCsv()
        {
            toolStripStatusLabel1.Text = "Saving";
            toolStripProgressBar1.Style = ProgressBarStyle.Continuous;
            toolStripProgressBar1.Visible = true;
            toolStripProgressBar1.Maximum = newDataTable.Rows.Count;
            toolStripProgressBar1.Step = 1;
            using (var dataTable = new DataTable())
            using (var reader = new DataTableReader(newDataTable))
            using (var stream = File.Create(SaveCsvPath, 4096, FileOptions.SequentialScan))
            using (var csv = new CsvWriter(new StreamWriter(stream), new Configuration() { Encoding = Encoding.UTF8,
                BufferSize = 4096, TrimOptions = TrimOptions.Trim | TrimOptions.InsideQuotes }))
            {
                try
                {
                    dataTable.Load(reader);
                    foreach (DataColumn column in dataTable.Columns)
                    {
                        csv.WriteField(column.ColumnName);
                    }
                    await csv.NextRecordAsync();
                    foreach (DataRow row in dataTable.Rows)
                    {
                        for (var i = 0; i < dataTable.Columns.Count; i++)
                        {
                            csv.WriteField(row[i]);
                        }
                        await csv.NextRecordAsync();
                        toolStripProgressBar1.PerformStep();
                    }
                    csvNeedsSave = false;
                    toolStripProgressBar1.Visible = false;
                    toolStripStatusLabel1.Text = "Ready";
                    return true;
                }
                catch (Exception) { }
            }
            toolStripProgressBar1.Visible = false;
            toolStripStatusLabel1.Text = "Ready";
            return false;
        }        
    }
}

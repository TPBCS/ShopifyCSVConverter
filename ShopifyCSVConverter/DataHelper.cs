﻿using CsvHelper;
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
    public class DataHelper
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

        public async Task<DataTable> BuildFromCsvParser()
        {
            char delimiter;
            using (var textReader = File.OpenText(Converter.OpenCsvPath))
            {
                delimiter = Detect(textReader, 5, ",|\t#.:;".ToCharArray());                
            }
            using (var reader = File.OpenText(Converter.OpenCsvPath))
            {
                using (var csv = new CsvParser(reader))
                {
                    csv.Configuration.Delimiter = delimiter.ToString();
                    csv.Configuration.BadDataFound = null;
                    string[] headers = await csv.ReadAsync();
                    if (headers != null && headers.Length > 0)
                    {
                        try
                        {
                            var table = new DataTable("OriginalData");
                            for (int i = 0; i < headers.Length; i++)
                            {
                                var header = headers[i];//"[ " + GetColumnName(i) + " ]\r\n" + 
                                table.Columns.Add(header);
                            }
                            while (true)
                            {
                                string[] record = await csv.ReadAsync();
                                if (record == null || record.Length == 0) break;
                                var row = table.NewRow();
                                for (int i = 0; i < record.Length; i++)
                                {
                                    row[i] = record[i];
                                }
                                table.Rows.Add(row);
                            }
                            return table;
                        }
                        catch (Exception exception)
                        {
                            #if DEBUG
                            MessageBox.Show($"{exception.Message} {exception.StackTrace}", $"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            #endif
                        }
                    }
                    return null;
                }
            }
        }

        private async Task<DataTable> buildFromCsvReader()
        {
            DataTable table = null;
            using (var reader = File.OpenText(Converter.OpenCsvPath))
            {
                using (var csv = new CsvReader(reader))
                {
                    if (csv.ReadHeader())
                    {
                        try
                        {
                            table = new DataTable("OrignalData");
                            var headers = csv.GetRecord<string[]>();
                            for (int i = 0; i < headers.Length; i++)
                            {
                                var header = "( " + GetColumnName(i) + " )" + headers[i];
                                table.Columns.Add(header);
                            }                            
                            while (await csv.ReadAsync())
                            {
                                var row = table.NewRow();
                                foreach (DataColumn column in table.Columns)
                                {
                                    row[column.ColumnName] = csv.GetField(column.DataType, column.ColumnName);
                                }
                                table.Rows.Add(row);
                            }
                            return table;
                        }
                        catch (Exception exception)
                        {
                            #if DEBUG
                            MessageBox.Show($"{exception.Message} {exception.StackTrace}", $"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            #endif
                        }
                    }
                    return null;
                }
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

        //        private async void LoadCsv()
        //        {
        //            CsvParser csv;
        //            int dataIndex = 0;

        //            try
        //            {
        //                using (StreamReader reader = File.OpenText(OpenCsvPath))
        //                {
        //                    csv = new CsvParser(reader);
        //                    csv.Configuration.TrimOptions = TrimOptions.Trim | TrimOptions.Trim;
        //                    csv.Read();
        //                    while (true)
        //                    {
        //                        string[] row = await csv.ReadAsync();
        //                        if (row == null || row.Length == 0) break;
        //                        originalData.Add(dataIndex, row);
        //                        dataIndex++;
        //                    }
        //                }
        //            }
        //            catch (Exception exception)
        //            {
        //#if DEBUG
        //                MessageBox.Show($"{exception.Message} {exception.StackTrace}", $"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //#endif
        //            }



        //        }
    }
}
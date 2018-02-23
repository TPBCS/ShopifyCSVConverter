using CsvHelper;
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
        private DataTable originalData;
        public DataTable OriginalData
        {
            get
            {
                if (originalData == null) buildFromCsvParser();
                return originalData;
            }
        }

        private async void buildFromCsvParser()
        {
            DataTable table = null;
            using (var reader = File.OpenText(Converter.OpenCsvPath))
            {
                using (var csv = new CsvParser(reader))
                {
                    string[] headers = await csv.ReadAsync();
                    if (headers != null && headers.Length > 0)
                    {
                        try
                        {
                            table = new DataTable("OriginalData");
                            foreach (string header in headers)
                            {
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
                        }
                        catch (Exception exception)
                        {
                            #if DEBUG
                            MessageBox.Show($"{exception.Message} {exception.StackTrace}", $"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            #endif
                        }
                    }
                }
            }
            originalData = table == null ? null : table;
        }

        private async void buildFromCsvReader()
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
                            foreach (string header in csv.GetRecord<string[]>())
                            {
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
                        }
                        catch (Exception exception)
                        {
                            #if DEBUG
                            MessageBox.Show($"{exception.Message} {exception.StackTrace}", $"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            #endif
                        }
                    }
                }
            }
            originalData = table == null ? null : table;
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

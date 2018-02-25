using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace ShopifyCSVConverter
{
    public partial class Converter
    {

        //save map
        private async void saveMapToolStripMenuItem_Click(object sender, EventArgs e)
        {
            await SaveCsvMap();
        }
        //load map
        private void loadMapToolStripMenuItem_Click(object sender, EventArgs e)
        {

            if (openCsvMapDialog.ShowDialog() == DialogResult.OK)
            {
                OpenCsvMapPath = openCsvMapDialog.FileName;
            }
            else return;

            List<string[]> parsedData = new List<string[]>();

            try
            {
                using (var stream = new FileStream(OpenCsvMapPath, FileMode.Open, FileAccess.Read))
                {
                    using (StreamReader reader = new StreamReader(stream))
                    {
                        var boxItems = reader.ReadLine().Split(new string[] {","}, StringSplitOptions.None);
                        
                        for (int i = 0; i < boxes.Length; i++)
                        {
                            boxes[i].SelectedIndex = hash45[boxItems[i]];
                        }
                    }                    
                }
            }
            catch (Exception)
            {
                
            }
        }
        //load original file
        private async void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult result = openCsvDialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                OpenCsvPath = openCsvDialog.FileName;
                toolStripStatusLabel1.Text = "Loading";
                toolStripProgressBar1.Style = ProgressBarStyle.Marquee;
                toolStripProgressBar1.Visible = true;
                originalDataTable = await LoadCsv();
                dataGridViewOriginal.DataSource = originalDataTable;
                FormatColumns(dataGridViewOriginal);
                toolStripProgressBar1.Visible = false;
                toolStripStatusLabel1.Text = "Ready";                  
            }
        }
        //convert
        private void convertToolStripMenuItem_Click(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Converting";
            toolStripProgressBar1.Style = ProgressBarStyle.Continuous;
            toolStripProgressBar1.Visible = true;
            toolStripProgressBar1.Maximum = originalDataTable.Rows.Count;
            toolStripProgressBar1.Step = 1;

            newDataTable = new DataTable();
            foreach (var column in shopifyColumns)
            {
                newDataTable.Columns.Add(column);
            }
            newDataTable.BeginLoadData();            
            foreach (var row in originalDataTable.Rows)
            {
                var rowItems = ((DataRow)row).ItemArray.Cast<string>().ToArray();
                var newRowItems = new string[boxes.Length];
                for (int i = 0; i < boxes.Length; i++)  
                {
                    if (boxes[i].SelectedIndex == 0) continue;
                    var key = boxes[i].GetItemText(boxes[i].SelectedItem);
                    newRowItems[i] = rowItems[hash100[key] - 1];
                }
                newDataTable.LoadDataRow(newRowItems, true);
                toolStripProgressBar1.PerformStep();
            }
            newDataTable.EndLoadData();
            dataGridViewConverted.DataSource = newDataTable;
            FormatColumns(dataGridViewConverted);            
            toolStripProgressBar1.Visible = false;
            toolStripStatusLabel1.Text = "Ready";
        }

        private void FormatColumns(DataGridView dataGridView)
        {
            for (int i = 0; i < dataGridView.Columns.Count; i++)
            {
                var column = dataGridView.Columns[i];
                var width = TextRenderer.MeasureText(column.HeaderText, dataGridView.ColumnHeadersDefaultCellStyle.Font).Width;
                column.Width = width > 200 ? 200 : width;
                column.MinimumWidth = 50;
            }
            var n = dataGridView.RowCount < 100 ? dataGridView.RowCount : 100;
            for (int i = 0; i < n; i++)
            {
                var cells = (dataGridView.Rows[i]).Cells;
                for (int j = 0; j < cells.Count; j++)
                {
                    var cell = (dataGridView.Rows[i]).Cells[j];
                    if (cell.Value == null || cell == null || cell.Value == DBNull.Value) continue;
                    var column = dataGridView.Columns[j];
                    var width = TextRenderer.MeasureText((string)cell.Value, cell.Style.Font).Width;
                    column.Width = width > 200 ? 200 : width > column.Width ? width : column.Width;
                    column.MinimumWidth = 50;
                    column.SortMode = DataGridViewColumnSortMode.NotSortable;
                }
            }
        }

        //Save converted file
        private async void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult result = saveCsvDialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                SaveCsvPath = saveCsvDialog.FileName;
                toolStripStatusLabel1.Text = "Saving";
                if(await SaveCsv()) toolStripStatusLabel1.Text = "Ready";
            }
        }
        //Exit menu item
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }
        //Ask to save on exit if files changed
        private void ShopifyCSVConverter_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (csvNeedsSave || csvMapNeedsSave)
            {
                var both = csvNeedsSave && csvMapNeedsSave;
                var fileOrFiles = both ? "files" : "file";
                var fileNames = both ? Path.GetFileName(OpenCsvPath) + "\r\n" + Path.GetFileName(OpenCsvMapPath)
                    : csvNeedsSave ? Path.GetFileName(OpenCsvPath) : Path.GetFileName(OpenCsvMapPath);
                DialogResult dialogResult = MessageBox.Show($"" +
                    $"Save changes to the following {fileOrFiles}?\n\n{fileNames}",
                    "Shopify CSV Converter", MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Question, MessageBoxDefaultButton.Button3);

                switch (dialogResult)
                {
                    case DialogResult.Cancel:
                        e.Cancel = true;
                        break;
                    case DialogResult.Yes:
                        Save();
                        break;
                    case DialogResult.No:
                        break;
                }
            }
        }

        //Display row numbers
        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            var dataGridView = sender as DataGridView;
            var rowIndex = (e.RowIndex + 1).ToString();
            var justifyRight = new StringFormat()
            {
                Alignment = StringAlignment.Far,
                LineAlignment = StringAlignment.Center
            };
            var headerBounds = new Rectangle(e.RowBounds.Left, e.RowBounds.Top, dataGridView.RowHeadersWidth - 10, e.RowBounds.Height);
            e.Graphics.DrawString(rowIndex, new Font("Microsoft Sans Serif", 8.25f, FontStyle.Bold), SystemBrushes.ControlText, headerBounds, justifyRight);
        }
        //Add letters to column headers, disable sorting
        private void dataGridView1_DataSourceChanged(object sender, EventArgs e)
        {
            DataGridView dataGridView = sender as DataGridView;
            DataTable table = new DataTable("Headers");
            for (int i = 0; i < dataGridView.ColumnCount; i++)
            {
                table.Columns.Add(GetColumnName(i));
            }
            if (dataGridView == dataGridViewOriginal)
            {
                dataGridViewOriginal.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
                dataGridViewLetterHeaders.DataSource = table;
                DisableColumnSorting(dataGridViewLetterHeaders);
                foreach (var column in dataGridViewLetterHeaders.Columns) ((DataGridViewColumn)column).MinimumWidth = 50;
                dataGridViewOriginal.ColumnWidthChanged += new DataGridViewColumnEventHandler(dataGridView1_ColumnWidthChanged);
                dataGridViewLetterHeaders.ColumnWidthChanged += new DataGridViewColumnEventHandler(dataGridView1_ColumnWidthChanged);
            }
            else if (dataGridView == dataGridViewConverted)
            {
                dataGridViewConverted.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
                dataGridViewConverted.ColumnWidthChanged += new DataGridViewColumnEventHandler(dataGridView1_ColumnWidthChanged);
            }
        }
        //Scroll Letter headers 
        private void dataGridView1_Scroll(object sender, ScrollEventArgs e)
        {
            DataGridView dataGridView = sender as DataGridView;
            if (e.ScrollOrientation == ScrollOrientation.HorizontalScroll)
            {
                if (dataGridView == dataGridViewOriginal)
                    dataGridViewLetterHeaders.HorizontalScrollingOffset = e.NewValue;
            }
        }
        //Match both header rows
        private void dataGridView1_ColumnWidthChanged(object sender, DataGridViewColumnEventArgs e)
        {
            DataGridView dataGridView = sender as DataGridView;
            if (dataGridView == dataGridViewOriginal)
                dataGridViewLetterHeaders.Columns[e.Column.Index].Width = e.Column.Width;
            else if (dataGridView == dataGridViewLetterHeaders)
                dataGridViewOriginal.Columns[e.Column.Index].Width = e.Column.Width;
            dataGridViewLetterHeaders.HorizontalScrollingOffset = dataGridViewOriginal.HorizontalScrollingOffset;
        }
        //Allow both row and column selection
        private void dataGridView_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            var dataGridView = sender as DataGridView;
            if (e.RowIndex == -1) dataGridView.SelectionMode = DataGridViewSelectionMode.ColumnHeaderSelect;
            if (e.ColumnIndex == -1) dataGridView.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect;
        }
    }
}
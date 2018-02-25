using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ShopifyCSVConverter
{
    public partial class Converter
    {

        //load map
        private void saveMapToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string path = "";

            if (saveCsvMapDialog.ShowDialog() == DialogResult.OK)
            {
                path = saveCsvMapDialog.FileName;
            }
            else return;

            string[] boxItems = new string[boxes.Length];

            for (int i = 0; i < boxes.Length; i++)
            {
                boxItems[i] = boxes[i].GetItemText(boxes[i].SelectedItem);
            }

            try
            {
                using (FileStream stream = new FileStream(path, FileMode.Create, FileAccess.Write))
                {
                    using (StreamWriter writer = new StreamWriter(stream))
                    {                        
                        writer.WriteLine(string.Join(",", boxItems));
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + ex.StackTrace);
            }
        }
        //load map
        private void loadMapToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string path = "";

            if (openCsvMapDialog.ShowDialog() == DialogResult.OK)
            {
                path = openCsvMapDialog.FileName;
            }
            else return;

            List<string[]> parsedData = new List<string[]>();

            try
            {
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read))
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
                dataGridView1.DataSource = await DataHelper.BuildFromCsvParser();
            }
        }
        //Save converted file
        private async void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult result = saveCsvDialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                SaveCsvPath = saveCsvDialog.FileName;
                await SaveCsv();
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
                table.Columns.Add(DataHelper.GetColumnName(i));
            }
            if (dataGridView == dataGridView1)
            {
                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
                dataGridView4.DataSource = table;
                DisableColumnSorting(dataGridView1);
                DisableColumnSorting(dataGridView4);
                dataGridView1.ColumnWidthChanged += new DataGridViewColumnEventHandler(dataGridView1_ColumnWidthChanged);
                dataGridView4.ColumnWidthChanged += new DataGridViewColumnEventHandler(dataGridView1_ColumnWidthChanged);
            }
            else if (dataGridView == dataGridView2)
            {
                dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
                dataGridView3.DataSource = table;
                DisableColumnSorting(dataGridView2);
                DisableColumnSorting(dataGridView3);
                dataGridView2.ColumnWidthChanged += new DataGridViewColumnEventHandler(dataGridView1_ColumnWidthChanged);
                dataGridView3.ColumnWidthChanged += new DataGridViewColumnEventHandler(dataGridView1_ColumnWidthChanged);
            }
        }
        //Scroll Letter headers 
        private void dataGridView1_Scroll(object sender, ScrollEventArgs e)
        {
            DataGridView dataGridView = sender as DataGridView;
            if (e.ScrollOrientation == ScrollOrientation.HorizontalScroll)
            {
                if (dataGridView == dataGridView1)
                    dataGridView4.HorizontalScrollingOffset = e.NewValue;
                else if (dataGridView == dataGridView2)
                    dataGridView3.HorizontalScrollingOffset = e.NewValue;
            }
        }
        //Match both header rows
        private void dataGridView1_ColumnWidthChanged(object sender, DataGridViewColumnEventArgs e)
        {
            DataGridView dataGridView = sender as DataGridView;
            if (dataGridView == dataGridView1)
                dataGridView4.Columns[e.Column.Index].Width = e.Column.Width;
            else if (dataGridView == dataGridView2)
                dataGridView3.Columns[e.Column.Index].Width = e.Column.Width;
            else if (dataGridView == dataGridView3)
                dataGridView2.Columns[e.Column.Index].Width = e.Column.Width;
            else if (dataGridView == dataGridView4)
                dataGridView1.Columns[e.Column.Index].Width = e.Column.Width;
        }
    }
}

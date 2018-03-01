using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Windows.Forms;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Linq;
using System.Globalization;

namespace ShopifyCSVConverter
{
    public partial class Converter : Form
    {
        internal static string OpenCsvPath;
        internal static string SaveCsvPath;
        internal static string OpenCsvMapPath;
        internal static string SaveCsvMapPath;
        private bool csvNeedsSave;
        private bool csvMapNeedsSave;
        private bool csvLoaded;
        private ComboBox[] boxes;
        private Dictionary<string, int> hash45;
        private Dictionary<string, int> hash100;
        private DataTable originalDataTable;
        private DataTable newDataTable;

        public Converter()
        {
            InitializeComponent();

            hash45 = GetHash45();

            hash100 = GetHash100();

            boxes = new ComboBox[]
            {
                mapComboBox1,
                mapComboBox2,
                mapComboBox3,
                mapComboBox4,
                mapComboBox5,
                mapComboBox6,
                mapComboBox7,
                mapComboBox8,
                mapComboBox9,
                mapComboBox10,
                mapComboBox11,
                mapComboBox12,
                mapComboBox13,
                mapComboBox14,
                mapComboBox15,
                mapComboBox16,
                mapComboBox17,
                mapComboBox18,
                mapComboBox19,
                mapComboBox20,
                mapComboBox21,
                mapComboBox22,
                mapComboBox23,
                mapComboBox24,
                mapComboBox25,
                mapComboBox26,
                mapComboBox27,
                mapComboBox28,
                mapComboBox29,
                mapComboBox30,
                mapComboBox31,
                mapComboBox32,
                mapComboBox33,
                mapComboBox34,
                mapComboBox35,
                mapComboBox36,
                mapComboBox37,
                mapComboBox38,
                mapComboBox39,
                mapComboBox40,
                mapComboBox41,
                mapComboBox42,
                mapComboBox43,
                mapComboBox44,
                mapComboBox45
            };

            foreach (var box in boxes)
            {
                box.DisplayMember = "Key";
                box.ValueMember = "Value";
                box.DataSource = new BindingSource(GetHash100(), null);
                EnableDoubleBuffering(box);
                box.SelectedValueChanged += ComboBox_ValueChanged;
            }

            var dataGridViews = new DataGridView[]
            {
                dataGridViewLetterHeaders,
                dataGridViewOriginal,
                dataGridViewConverted
            };

            foreach (var dataGridView in dataGridViews)
            {
                if (dataGridView == dataGridViewOriginal || dataGridView == dataGridViewConverted) dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.ColumnHeader;                
                EnableDoubleBuffering(dataGridView);
            }

            UpdateBoxes();

            convertToolStripMenuItem.Enabled = CanConvert();
            saveToolStripMenuItem.Enabled = false;
            toolStripProgressBar1.Visible = false;
        }        

        protected override void OnCreateControl()
        {
            base.OnCreateControl();
            this.SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.CacheText, true);
        }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= WS_EX_COMPOSITED;
                return cp;
            }
        }

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);
            if (m.Msg == 0x0112)
            {
                int wParam = (m.WParam.ToInt32() & 0xFFF0);// Maximize by double-clicking title bar
                if (wParam == 0xF030 // Maximize event - SC_MAXIMIZE from Winuser.h
                    || wParam == 0xF120 // Restore event - SC_RESTORE from Winuser.h
                    || wParam == 0XF020) // Minimize event - SC_MINIMIZE from Winuser.h
                    UpdateBoxes();
            }
        }

        public void BeginUpdate()
        {
            NativeMethods.SendMessage(this.Handle, WM_SETREDRAW, 0, 0);
        }

        public void EndUpdate()
        {
            NativeMethods.SendMessage(this.Handle, WM_SETREDRAW, 1, 0);
            Invalidate(true);
        }
                
        private void EnableDoubleBuffering(object target)
        {
            if (!SystemInformation.TerminalServerSession)
            {
                Type dgvType = dataGridViewOriginal.GetType();
                PropertyInfo pi = dgvType.GetProperty("DoubleBuffered",
                  BindingFlags.Instance | BindingFlags.NonPublic);
                pi.SetValue(target, true, null);
            }
        }

        private void UpdateBoxes()
        {
            foreach (var box in boxes)
            {
                NativeMethods.SendMessage(box.Handle, CB_SETITEMHEIGHT, -1, headerLabel1.Height - 6);
                box.Refresh();
            }
        }             

        private void DisableColumnSorting(DataGridView dataGridView)
        {
            foreach (var column in dataGridView.Columns)
            {
                var col = column as DataGridViewColumn;
                col.SortMode = DataGridViewColumnSortMode.NotSortable; 
            }
        }

        private void Save()
        {
            SaveCsvPath = Path.Combine(Application.StartupPath, "converted.csv");
            SaveCsvMapPath = Path.Combine(Application.StartupPath, "csv-map.sccm");
            toolStripStatusLabel1.Text = "Saving";
            if (csvNeedsSave && csvMapNeedsSave)
            {
                SaveCsvMap();
                SaveCsv();
            }
            else if (csvMapNeedsSave) SaveCsvMap(); 
            else if (csvNeedsSave) SaveCsv();
            toolStripStatusLabel1.Text = "Ready";
        }

        private void SaveCsvMap()
        {
            csvMapNeedsSave = false;                        

            if ((saveCsvMapDialog.ShowDialog() == DialogResult.OK))
            {
                SaveCsvMapPath = saveCsvMapDialog.FileName;

                string[] boxItems = new string[boxes.Length];

                for (int i = 0; i < boxes.Length; i++)
                {
                    boxItems[i] = boxes[i].GetItemText(boxes[i].SelectedItem);
                }
                
                FileStream stream = null;
                try
                {
                    stream = new FileStream(SaveCsvMapPath, FileMode.Create, FileAccess.Write);                    
                    using (StreamWriter writer = new StreamWriter(stream))
                    {
                        stream = null;
                        writer.WriteLine(string.Join(",", boxItems));
                    }
                }
                finally
                {
                    if (stream != null) stream.Dispose();
                }                
            } 
        }
        private bool CanConvert()
        {
            var anyBoxLoaded = false;
            foreach (var box in boxes) if (box.SelectedIndex > 0) anyBoxLoaded = true;
            return anyBoxLoaded && csvLoaded;
        }
        //Capitalize Words
        private string CapitalizeWords(string sentenceOrWord)
        {
            if (string.IsNullOrWhiteSpace(sentenceOrWord)) return sentenceOrWord;
            var stringArray = sentenceOrWord.Split(char.Parse(" "));
            for (int i = 0; i < stringArray.Length; i++)
            {
                stringArray[i] = CultureInfo.InvariantCulture.TextInfo.ToTitleCase(stringArray[i].ToLower());
            }
            return string.Join(" ", stringArray);
        }
        //Format Tags
        private string FormatTags(string tags)
        {
            tags = tags.Replace("&", "And");
            var tagsArray = tags.Split(char.Parse(","));
            var newTagsArray = new List<string>();
            foreach (var tag in tagsArray)
            {
                var tagArray = tag.Split(char.Parse(" "));
                var newTagArray = new List<string>(); 
                foreach (var str in tagArray)
                {
                    newTagArray.Add(str.ToLower().Trim().First().ToString().ToUpper() + str.Remove(0, 1));
                }
                var newTag = string.Join("", newTagArray);
                if (!newTagsArray.Contains(newTag) && newTag.ToLower() != "allclearance") newTagsArray.Add(newTag);
            }            
            return string.Join(",", newTagsArray);
        }
        //Format Handle
        private string FormatHandle(string handle)
        {
            var illegalChars = handle.Trim().ToCharArray().ToList().FindAll(c => !char.IsLetter(c) && c != char.Parse("-"));
            foreach (var illegalChar in illegalChars)
            {
                handle = handle.Replace(illegalChar, char.Parse("-"));
            }
            return handle.Replace("---", "-").Replace("--", "-").ToLower();
        }
        //Format Image ALT Tag
        private string FormatAltTag(string altTagHandle, string altTagDescription)
        {  
            return $"Image of product {altTagHandle}. {altTagDescription}.";
        }
        //Set column width to column content width or 200, whichever is less
        private void FormatColumns(DataGridView dataGridView)
        {
            for (int i = 0; i < dataGridView.Columns.Count; i++)
            {
                var column = dataGridView.Columns[i];
                var width = TextRenderer.MeasureText(column.HeaderText, dataGridView.ColumnHeadersDefaultCellStyle.Font).Width;
                column.MinimumWidth = 50;
                column.Width = width > 200 ? 200 : width;
            }
            var n = dataGridView.RowCount < 100 ? dataGridView.RowCount : 100;
            for (int i = 0; i < n; i++)
            {
                try
                {
                    var cells = (dataGridView.Rows[i]).Cells;
                    for (int j = 0; j < cells.Count; j++)
                    {
                        var column = dataGridView.Columns[j];
                        column.SortMode = DataGridViewColumnSortMode.NotSortable;
                        var cell = (dataGridView.Rows[i]).Cells[j];
                        if (cell.Value == null || cell == null || cell.Value == DBNull.Value) continue;
                        var width = TextRenderer.MeasureText((string)cell.Value, cell.Style.Font).Width;
                        column.MinimumWidth = 50;
                        column.Width = width > 200 ? 200 : width > column.Width ? width : column.Width;
                    }
                }
                catch (Exception) { }
            }
        }
    }

    internal static class NativeMethods
    {
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        internal static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

        internal static void SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam)
        {
            SendMessage(hWnd, Msg, (IntPtr)wParam, (IntPtr)lParam);
        }
    }
}
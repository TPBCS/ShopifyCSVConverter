using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using CsvHelper;
using CsvHelper.Configuration;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ShopifyCSVConverter
{
    public partial class Converter : Form
    {
        #region Fields
        internal static string OpenCsvPath;
        internal static string SaveCsvPath;
        internal static string OpenCsvMapPath;
        internal static string SaveCsvMapPath;
        private bool csvNeedsSave;
        private bool csvMapNeedsSave;
        private Font mapFont = new Font("Arial", 10, FontStyle.Bold);
        private ComboBox[] boxes;
        private Dictionary<string, int> hash45;
        private Dictionary<string, int> hash100;
        private DataHelper dataHelper;
        private DataHelper DataHelper => dataHelper != null ? dataHelper : dataHelper = new DataHelper();
        #endregion

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
                cp.ExStyle |= NativeMethods.WS_EX_COMPOSITED;
                return cp;
            }
        }

        public void BeginUpdate()
        {
            NativeMethods.SendMessage(this.Handle, NativeMethods.WM_SETREDRAW, IntPtr.Zero, IntPtr.Zero);
        }

        public void EndUpdate()
        {
            NativeMethods.SendMessage(this.Handle, NativeMethods.WM_SETREDRAW, new IntPtr(1), IntPtr.Zero);
            Parent.Invalidate(true);
        }

        public Converter()
        {
            InitializeComponent();            
            hash45 = getHash45();
            hash100 = getHash100();

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
                box.DataSource = new BindingSource(getHash100(), null);
            }

            updateBoxes();
        }

        #region ComboBox        

        private void SetComboBoxHeight(IntPtr comboBoxHandle, Int32 comboBoxDesiredHeight)
        {
            SendMessage(comboBoxHandle, CB_SETITEMHEIGHT, -1, comboBoxDesiredHeight);
        }

        public const Int32 CB_SETITEMHEIGHT = 0x153;

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, Int32 wParam, Int32 lParam);

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);

            
            if (m.Msg == 0x0112)
            {
                int wParam = (m.WParam.ToInt32() & 0xFFF0);// Maximize by double-clicking title bar
                if (wParam == 0xF030 // Maximize event - SC_MAXIMIZE from Winuser.h
                    || wParam == 0xF120 // Restore event - SC_RESTORE from Winuser.h
                    || wParam == 0XF020) // Minimize event - SC_MINIMIZE from Winuser.h
                    updateBoxes();                
            }
        }

        private void updateBoxes()
        {
            foreach (var box in boxes)
            {
                SetComboBoxHeight(box.Handle, headerLabel1.Height - 6);
                box.Refresh();
            }            
        }        

        private Dictionary<string, int> getHash45()
        {
            return new Dictionary<string, int>()
            {
                [""] = 0,
                ["A"] = 1,
                ["B"] = 2,
                ["C"] = 3,
                ["D"] = 4,
                ["E"] = 5,
                ["F"] = 6,
                ["G"] = 7,
                ["H"] = 8,
                ["I"] = 9,
                ["J"] = 10,
                ["K"] = 11,
                ["L"] = 12,
                ["M"] = 13,
                ["N"] = 14,
                ["O"] = 15,
                ["P"] = 16,
                ["Q"] = 17,
                ["R"] = 18,
                ["S"] = 19,
                ["T"] = 20,
                ["U"] = 21,
                ["V"] = 22,
                ["W"] = 23,
                ["X"] = 24,
                ["Y"] = 25,
                ["Z"] = 26,
                ["AA"] = 27,
                ["AB"] = 28,
                ["AC"] = 29,
                ["AD"] = 30,
                ["AE"] = 31,
                ["AF"] = 32,
                ["AG"] = 33,
                ["AH"] = 34,
                ["AI"] = 35,
                ["AJ"] = 36,
                ["AK"] = 37,
                ["AL"] = 38,
                ["AM"] = 39,
                ["AN"] = 40,
                ["AO"] = 41,
                ["AP"] = 42,
                ["AQ"] = 43,
                ["AR"] = 44,
                ["AS"] = 45,
                ["AT"] = 46
            };
        }

        internal static Dictionary<string, int> getHash100()
        {
            return new Dictionary<string, int>()
            {
                [""] = 0,
                ["A"] = 1,
                ["B"] = 2,
                ["C"] = 3,
                ["D"] = 4,
                ["E"] = 5,
                ["F"] = 6,
                ["G"] = 7,
                ["H"] = 8,
                ["I"] = 9,
                ["J"] = 10,
                ["K"] = 11,
                ["L"] = 12,
                ["M"] = 13,
                ["N"] = 14,
                ["O"] = 15,
                ["P"] = 16,
                ["Q"] = 17,
                ["R"] = 18,
                ["S"] = 19,
                ["T"] = 20,
                ["U"] = 21,
                ["V"] = 22,
                ["W"] = 23,
                ["X"] = 24,
                ["Y"] = 25,
                ["Z"] = 26,
                ["AA"] = 27,
                ["AB"] = 28,
                ["AC"] = 29,
                ["AD"] = 30,
                ["AE"] = 31,
                ["AF"] = 32,
                ["AG"] = 33,
                ["AH"] = 34,
                ["AI"] = 35,
                ["AJ"] = 36,
                ["AK"] = 37,
                ["AL"] = 38,
                ["AM"] = 39,
                ["AN"] = 40,
                ["AO"] = 41,
                ["AP"] = 42,
                ["AQ"] = 43,
                ["AR"] = 44,
                ["AS"] = 45,
                ["AT"] = 46,
                ["AU"] = 47,
                ["AV"] = 48,
                ["AW"] = 49,
                ["AX"] = 50,
                ["AY"] = 51,
                ["AZ"] = 52,
                ["BA"] = 53,
                ["BB"] = 54,
                ["BC"] = 55,
                ["BD"] = 56,
                ["BE"] = 57,
                ["BF"] = 58,
                ["BG"] = 59,
                ["BH"] = 60,
                ["BI"] = 61,
                ["BJ"] = 62,
                ["BK"] = 63,
                ["BL"] = 64,
                ["BM"] = 65,
                ["BN"] = 66,
                ["BO"] = 67,
                ["BP"] = 68,
                ["BQ"] = 69,
                ["BR"] = 70,
                ["BS"] = 71,
                ["BT"] = 72,
                ["BU"] = 73,
                ["BV"] = 74,
                ["BW"] = 75,
                ["BX"] = 76,
                ["BY"] = 77,
                ["BZ"] = 78,
                ["CA"] = 79,
                ["CB"] = 80,
                ["CC"] = 81,
                ["CD"] = 82,
                ["CE"] = 83,
                ["CF"] = 84,
                ["CG"] = 85,
                ["CH"] = 86,
                ["CI"] = 87,
                ["CJ"] = 88,
                ["CK"] = 89,
                ["CL"] = 90,
                ["CM"] = 91,
                ["CN"] = 92,
                ["CO"] = 93,
                ["CP"] = 94,
                ["CQ"] = 95,
                ["CR"] = 96,
                ["CS"] = 97,
                ["CT"] = 98,
                ["CU"] = 99,
                ["CV"] = 100
            };
        }
        #endregion
        #region Menu
        private async void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult result = openCsvDialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                OpenCsvPath = openCsvDialog.FileName;
                dataGridView1.DataSource = await DataHelper.BuildFromCsvParser();

            }
        }        

        private async void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult result = saveCsvDialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                SaveCsvPath = saveCsvDialog.FileName;
                await SaveCsv();
            }
        }

        private async Task<bool> SaveCsv()
        {
            return true;
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }
        #endregion

        #region PromptForSave
        private void ShopifyCSVConverter_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (csvNeedsSave || csvMapNeedsSave)
            {                
                var both = csvNeedsSave && csvMapNeedsSave;
                var fileOrFiles = both ? "files" : "file";
                var fileNames = both ? Path.GetFileName(OpenCsvPath) + "\r\n" + Path.GetFileName(OpenCsvMapPath) : csvNeedsSave ? Path.GetFileName(OpenCsvPath) : Path.GetFileName(OpenCsvMapPath);
                DialogResult dialogResult = MessageBox.Show($"Save changes to the following {fileOrFiles}?\n\n{fileNames}", "Shopify CSV Converter", MessageBoxButtons.YesNoCancel, 
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

        private void Save()
        {
            throw new NotImplementedException();
        }
        #endregion

        private void Converter_DragDrop(object sender, DragEventArgs e)
        {
            
        }

        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            var grid = sender as DataGridView;
            var rowIndex = (e.RowIndex + 1).ToString() + " ";

            var rightFormat = new StringFormat()
            {
                // right alignment might actually make more sense for numbers
                Alignment = StringAlignment.Far,
                LineAlignment = StringAlignment.Center
            };

            var headerBounds = new Rectangle(e.RowBounds.Left, e.RowBounds.Top, grid.RowHeadersWidth, e.RowBounds.Height);
            e.Graphics.DrawString(rowIndex, this.Font, SystemBrushes.ControlText, headerBounds, rightFormat);
        }

        private void dataGridView1_DataSourceChanged(object sender, EventArgs e)
        {
            DataGridView dataGridView = sender as DataGridView;
            DataTable table = new DataTable("Headers");
            for (int i = 0; i < dataGridView.ColumnCount; i++)
            {
                table.Columns.Add(DataHelper.GetColumnName(i));
            }
            if(dataGridView == dataGridView1)
            {
                dataGridView4.DataSource = table;
                dataGridView4.ColumnWidthChanged += new DataGridViewColumnEventHandler(dataGridView1_ColumnWidthChanged);
            }                
            else if (dataGridView == dataGridView2)
            {
                dataGridView3.DataSource = table;
                dataGridView3.ColumnWidthChanged += new DataGridViewColumnEventHandler(dataGridView1_ColumnWidthChanged);
            }
        }        

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

        private void dataGridView1_ColumnWidthChanged(object sender, DataGridViewColumnEventArgs e)
        {
            DataGridView dataGridView = sender as DataGridView;
            if (dataGridView == dataGridView1)
                dataGridView4.Columns[e.Column.Index].Width = e.Column.Width;
            else if (dataGridView == dataGridView2)
                dataGridView3.Columns[e.Column.Index].Width = e.Column.Width;
        }
    }

    public static class NativeMethods
    {
        public static int WM_SETREDRAW = 0x000B; //uint WM_SETREDRAW
        public static int WS_EX_COMPOSITED = 0x02000000;       

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam); 
    }
}

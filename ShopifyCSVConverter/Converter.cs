using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;
using System.Runtime.InteropServices;

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
        private ComboBox[] boxes;
        private Dictionary<string, int> hash45;
        private Dictionary<string, int> hash100;
        private DataTable originalDataTable;
        private DataTable newDataTable;

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
                EnableDoubleBuffering(box);
                box.SelectedValueChanged += ComboBox_ValueChanged;
            }

            var dataGridViews = new DataGridView[]
            {
                dataGridViewLetterHeaders,
                dataGridViewOriginal,
                dataGridViewConverted
            };
            //gets called when data is not yet bound so cant deal with columns and rows here
            foreach (var dataGridView in dataGridViews)
            {
                if (dataGridView == dataGridViewOriginal || dataGridView == dataGridViewConverted) dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.ColumnHeader;                
                EnableDoubleBuffering(dataGridView);
            }

            UpdateBoxes();
            
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
            NativeMethods.SendMessage(this.Handle, WM_SETREDRAW, false, 0);
        }

        public void EndUpdate()
        {
            NativeMethods.SendMessage(this.Handle, WM_SETREDRAW, true, 0);
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


        private async void Save()
        {
            toolStripStatusLabel1.Text = "Saving";
            if (csvNeedsSave && csvMapNeedsSave)
            {
                await SaveCsv();
                await SaveCsvMap();
            }
            else if (csvMapNeedsSave)
            {
                await SaveCsvMap();
            }
            else if(csvNeedsSave)
            {
                await SaveCsv();
            }
            toolStripStatusLabel1.Text = "Ready";
        }

        private async Task<bool> SaveCsvMap()
        {
            csvMapNeedsSave = false;                        

            if (!(saveCsvMapDialog.ShowDialog() == DialogResult.OK)) return false;

            SaveCsvMapPath = saveCsvMapDialog.FileName;            

            string[] boxItems = new string[boxes.Length];

            for (int i = 0; i < boxes.Length; i++)
            {
                boxItems[i] = boxes[i].GetItemText(boxes[i].SelectedItem);
            }

            try
            {
                using (FileStream stream = new FileStream(SaveCsvMapPath, FileMode.Create, FileAccess.Write))
                {
                    using (StreamWriter writer = new StreamWriter(stream))
                    {
                        await writer.WriteLineAsync(string.Join(",", boxItems));
                    }
                }
                return true;
            }
            catch (Exception) { }
            return false;
        }
    }

    public static class NativeMethods
    {
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, bool wParam, Int32 lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, Int32 wParam, Int32 lParam);
    }
}

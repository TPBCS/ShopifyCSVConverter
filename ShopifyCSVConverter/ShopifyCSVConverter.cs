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
        private bool csvNeedsSave;
        private bool csvMapNeedsSave;
        private string openCsvPath;
        private string openCsvMapPath;
        private string saveCsvPath;
        private string saveCsvMapPath;
        private Font mapFont = new Font("Arial", 10, FontStyle.Bold);
        private ComboBox[] boxes;
        private Dictionary<string, int> hash45;
        private Dictionary<string, int> hash100;
        private Dictionary<int, string[]> originalData;        
        #endregion

        public Converter()
        {
            InitializeComponent();

            hash45 = getHash45();
            hash100 = getHash100();

            boxes = new ComboBox[]
            {
                comboBox1,
                //comboBox3,
                //comboBox4,
                //comboBox5,
                //comboBox6,
                //comboBox7,
                //comboBox8,
                //comboBox9,
                //comboBox10,
                //comboBox11,
                //comboBox12,
                //comboBox13,
                //comboBox14,
                //comboBox15,
                //comboBox16,
                //comboBox17,
                //comboBox18,
                //comboBox19,
                //comboBox20,
                //comboBox21,
                //comboBox22,
                //comboBox23,
                //comboBox24,
                //comboBox25,
                //comboBox26,
                //comboBox27,
                //comboBox28,
                //comboBox29,
                //comboBox30,
                //comboBox31,
                //comboBox32,
                //comboBox33,
                //comboBox34,
                //comboBox35,
                //comboBox36,
                //comboBox37,
                //comboBox38,
                //comboBox39,
                //comboBox40,
                //comboBox41,
                //comboBox42,
                //comboBox43,
                //comboBox44,
                //comboBox45
            };

            foreach (var box in boxes)
            {
                box.DisplayMember = "Key";
                box.ValueMember = "Value";
                box.DataSource = new BindingSource(getHash100(), null);
            }

            updateUI();
        }
        #region ComboBox
        private void comboBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            var box = sender as ComboBox;
            Brush foreground = new SolidBrush(e.ForeColor);
            Brush background = new SolidBrush(Color.White);
            var item = ((KeyValuePair<string, int>)box.Items[e.Index]).Key;
            StringFormat stringFormat = new StringFormat();
            stringFormat.LineAlignment = StringAlignment.Center;
            stringFormat.Alignment = StringAlignment.Center;

            if((e.State & DrawItemState.Selected) == DrawItemState.Selected)
            {
                foreground = Brushes.White;
                background = new SolidBrush(Color.Turquoise);
            };
            e.Graphics.FillRectangle(background, e.Bounds);
            e.Graphics.DrawString(item, box.Font, foreground, e.Bounds, stringFormat); 
        }

        [DllImport("user32.dll")]
        static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, Int32 wParam, Int32 lParam);

        private const Int32 CB_SETITEMHEIGHT = 0x153;

        private void SetComboBoxHeight(IntPtr comboBoxHandle, Int32 comboBoxDesiredHeight)
        {
            SendMessage(comboBoxHandle, CB_SETITEMHEIGHT, -1, comboBoxDesiredHeight);
        }

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);

            // WM_SYSCOMMAND
            if (m.Msg == 0x0112)
            {
                int wParam = (m.WParam.ToInt32() & 0xFFF0);// Add double click titlebar support.
                if (wParam == 0xF030 // Maximize event - SC_MAXIMIZE from Winuser.h
                    || wParam == 0xF120 // Restore event - SC_RESTORE from Winuser.h
                     || wParam == 0XF020) // Minimize event - SC_MINIMIZE from Winuser.h
                
                {
                    updateUI();
                }
            }
        }

        private void updateUI()
        {
            SetComboBoxHeight(comboBox1.Handle, label1.Height - 6);
            comboBox1.Refresh();
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
        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult result = openCsvDialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                openCsvPath = openCsvDialog.FileName;
                LoadCsv();
            }
        }

        private async void LoadCsv()
        {
            CsvParser csv;
            originalData = new Dictionary<int, string[]>();
            int dataIndex = 0;

            try
            {
                using (StreamReader reader = File.OpenText(openCsvPath))
                {                    
                    csv = new CsvParser(reader);
                    csv.Configuration.TrimOptions = TrimOptions.Trim | TrimOptions.Trim;
                    csv.Read();
                    while (true)
                    {
                        string[] row = await csv.ReadAsync();
                        if (row == null || row.Length == 0) break;
                        originalData.Add(dataIndex, row);
                        dataIndex++;                        
                    }
                }
            }
            catch (Exception exception)
            {
                #if DEBUG
                MessageBox.Show($"{exception.Message} {exception.StackTrace}", $"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                #endif
            }
            
            
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult result = saveCsvDialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                saveCsvPath = saveCsvDialog.FileName;
                SaveCsv();
            }
        }

        private async void SaveCsv()
        {
            
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
                var fileNames = both ? Path.GetFileName(openCsvPath) + "\r\n" + Path.GetFileName(openCsvMapPath) : csvNeedsSave ? Path.GetFileName(openCsvPath) : Path.GetFileName(openCsvMapPath);
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
    }
}

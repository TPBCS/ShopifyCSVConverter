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

namespace ShopifyCSVConverter
{
    public partial class ShopifyCSVConverter : Form
    {
        #region Fields
        private bool csvNeedsSave;
        private bool csvMapNeedsSave;
        private string openCsvPath;
        private string openCsvMapPath;
        private string saveCsvPath;
        private string saveCsvMapPath;
        #endregion

        public ShopifyCSVConverter()
        {
            InitializeComponent();
        }

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
            throw new NotImplementedException();
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
            throw new NotImplementedException();
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

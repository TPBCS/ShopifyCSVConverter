using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ShopifyCSVConverter
{
    public partial class ShopifyCSVConverter : Form
    {
        private bool needsSave;

        public ShopifyCSVConverter()
        {
            InitializeComponent();
        }

        #region Menu
        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }
        #endregion

        #region Form
        private void ShopifyCSVConverter_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (needsSave)
            {
                DialogResult dialogResult = MessageBox.Show("Would you like to save all changes?", "Shopify CSV Converter", MessageBoxButtons.YesNoCancel, 
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

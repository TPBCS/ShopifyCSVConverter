using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace ShopifyCSVConverter
{    
    class HeaderLabel : Label
    {
        public HeaderLabel() : base()
        {
            this.AutoSize = true;
            this.BackColor = Color.FromArgb(((int)(((byte)(240)))), ((int)(((byte)(240)))), ((int)(((byte)(240)))));
            this.BorderStyle = BorderStyle.FixedSingle;
            this.Dock = DockStyle.Fill;
            this.Font = new Font("Arial", 10F, FontStyle.Bold, GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new Padding(0);
            this.TextAlign = ContentAlignment.MiddleCenter;
        }            
    }

    class MapComboBox : ComboBox
    {
        public MapComboBox() : base()
        {
            this.Anchor = ((AnchorStyles)((((AnchorStyles.Top | AnchorStyles.Bottom)
            | AnchorStyles.Left)
            | AnchorStyles.Right)));
            this.DrawMode = DrawMode.OwnerDrawFixed;
            this.DropDownStyle = ComboBoxStyle.DropDownList; this.Font = new Font("Arial", 10F, FontStyle.Bold, GraphicsUnit.Point, ((byte)(0)));
            this.FormattingEnabled = true;
            this.ItemHeight = 21;
            this.Margin = new Padding(1, 0, 0, 0);
            this.Size = new Size(49, 27);
            this.DrawItem += new DrawItemEventHandler(this.comboBox_DrawItem);
            this.DropDownHeight = 300;
            this.IntegralHeight = false;
        }

        private void comboBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            var box = sender as ComboBox;
            if (e.Index >= 0)
            {
                Brush foreground = (e.State & DrawItemState.Selected) == DrawItemState.Selected ? Brushes.White : new SolidBrush(e.ForeColor);
                Brush background = (e.State & DrawItemState.Selected) == DrawItemState.Selected ? new SolidBrush(Color.Turquoise) : new SolidBrush(Color.White);                
                var item = ((KeyValuePair<string, int>)box.Items[e.Index]).Key;
                StringFormat stringFormat = new StringFormat();
                stringFormat.LineAlignment = StringAlignment.Center;
                stringFormat.Alignment = StringAlignment.Center;
                e.Graphics.FillRectangle(background, e.Bounds);
                e.Graphics.DrawString(item, box.Font, foreground, e.Bounds, stringFormat);
            }            
        }
    }

    class HorizontalDividerLabel : Label
    {
        public HorizontalDividerLabel() : base()
        {
            this.AutoSize = false;
            this.Anchor = ((AnchorStyles)((AnchorStyles.Left | AnchorStyles.Right)));
            this.BorderStyle = BorderStyle.Fixed3D;            
            this.Margin = new Padding(1, 1, 0, 0);
            this.Height = 2;
        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AcElectricalSchemePlugin
{
    public partial class ChangeAttributeForm : Form
    {
        public ChangeAttributeForm()
        {
            InitializeComponent();
        }

        private void okBtn_Click(object sender, EventArgs e)
        {
            if (attrListBox.SelectedIndex>=0)
            {
                ChangeAttributeClass.changeAttributes(attrListBox.Items[attrListBox.SelectedIndex].ToString(), attrNameListBox.Text);
                this.Close();
            }
        }
    }
}

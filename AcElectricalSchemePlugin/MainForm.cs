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
    public partial class MainForm : MetroFramework.Forms.MetroForm
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void ConnectionSchemeBtn_Click(object sender, EventArgs e)
        {
            AcPlugin.ConnectionScheme();
        }

        private void ContourSchemeBtn_Click(object sender, EventArgs e)
        {
            AcPlugin.ContourScheme();
        }

        private void ControlSchemeToTxtBtn_Click(object sender, EventArgs e)
        {
            AcPlugin.ParseControlSchemeToTxt();
        }

        private void MarkBtn_Click(object sender, EventArgs e)
        {
            AcPlugin.Mark();
        }

        private void ControlSchemeBtn_Click(object sender, EventArgs e)
        {
            AcPlugin.ControlScheme();
        }

        private void TerminalsBtn_Click(object sender, EventArgs e)
        {
            AcPlugin.TerminalsWeight();
        }

        private void ChangeArticlesBtn_Click(object sender, EventArgs e)
        {
            AcPlugin.ChangeArticles();
        }

        private void SLineSchemeBtn_Click(object sender, EventArgs e)
        {
            AcPlugin.SLineScheme();
        }

        private void ChangeAttributesBtn_Click(object sender, EventArgs e)
        {
            AcPlugin.ChangeAttributes();
        }

        private void ChangeTextHeghtBtn_Click(object sender, EventArgs e)
        {
            AcPlugin.ChangeTextHeght();
        }

        private void MarkExportBtn_Click(object sender, EventArgs e)
        {
            AcPlugin.MarkExport();
        }

        private void HelpBtn_Click(object sender, EventArgs e)
        {
            HelpForm hForm = new HelpForm();
            hForm.Show();
        }
    }
}

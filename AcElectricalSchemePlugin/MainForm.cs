using System;

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
            this.Visible = false;
            AcPlugin.ConnectionScheme();
            this.Visible = true;
        }

        private void ContourSchemeBtn_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            AcPlugin.ContourScheme();
            this.Visible = true;
        }

        private void ControlSchemeToTxtBtn_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            AcPlugin.ParseControlSchemeToTxt();
            this.Visible = true;
        }

        private void MarkBtn_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            AcPlugin.Mark();
            this.Visible = true;
        }

        private void ControlSchemeBtn_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            AcPlugin.ControlScheme();
            this.Visible = true;
        }

        private void TerminalsBtn_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            AcPlugin.TerminalsWeight();
            this.Visible = true;
        }

        private void ChangeArticlesBtn_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            AcPlugin.ChangeArticles();
            this.Visible = true;
        }

        private void SLineSchemeBtn_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            AcPlugin.SLineScheme();
            this.Visible = true;
        }

        private void ChangeAttributesBtn_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            AcPlugin.ChangeAttributes();
            this.Visible = true;
        }

        private void ChangeTextHeghtBtn_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            AcPlugin.ChangeTextHeght();
            this.Visible = true;
        }

        private void MarkExportBtn_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            AcPlugin.MarkExport();
            this.Visible = true;
        }

        private void HelpBtn_Click(object sender, EventArgs e)
        {
            HelpForm hForm = new HelpForm();
            hForm.Show();
        }
    }
}

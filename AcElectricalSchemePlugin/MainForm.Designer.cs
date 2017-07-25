namespace AcElectricalSchemePlugin
{
    partial class MainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.ConnectionSchemeBtn = new MetroFramework.Controls.MetroButton();
            this.ContourSchemeBtn = new MetroFramework.Controls.MetroButton();
            this.MarkBtn = new MetroFramework.Controls.MetroButton();
            this.ControlSchemeBtn = new MetroFramework.Controls.MetroButton();
            this.TerminalsBtn = new MetroFramework.Controls.MetroButton();
            this.ChangeArticlesBtn = new MetroFramework.Controls.MetroButton();
            this.SLineSchemeBtn = new MetroFramework.Controls.MetroButton();
            this.ChangeAttributesBtn = new MetroFramework.Controls.MetroButton();
            this.ChangeTextHeghtBtn = new MetroFramework.Controls.MetroButton();
            this.MarkExportBtn = new MetroFramework.Controls.MetroButton();
            this.HelpBtn = new MetroFramework.Controls.MetroButton();
            this.ControlSchemeToTxtBtn = new MetroFramework.Controls.MetroButton();
            this.SuspendLayout();
            // 
            // ConnectionSchemeBtn
            // 
            this.ConnectionSchemeBtn.Location = new System.Drawing.Point(137, 107);
            this.ConnectionSchemeBtn.Name = "ConnectionSchemeBtn";
            this.ConnectionSchemeBtn.Size = new System.Drawing.Size(198, 23);
            this.ConnectionSchemeBtn.Style = MetroFramework.MetroColorStyle.Yellow;
            this.ConnectionSchemeBtn.TabIndex = 0;
            this.ConnectionSchemeBtn.Text = "Схема внешних подключений";
            this.ConnectionSchemeBtn.Theme = MetroFramework.MetroThemeStyle.Dark;
            this.ConnectionSchemeBtn.Click += new System.EventHandler(this.ConnectionSchemeBtn_Click);
            // 
            // ContourSchemeBtn
            // 
            this.ContourSchemeBtn.Location = new System.Drawing.Point(247, 136);
            this.ContourSchemeBtn.Name = "ContourSchemeBtn";
            this.ContourSchemeBtn.Size = new System.Drawing.Size(123, 23);
            this.ContourSchemeBtn.Style = MetroFramework.MetroColorStyle.Yellow;
            this.ContourSchemeBtn.TabIndex = 1;
            this.ContourSchemeBtn.Text = "Схема контуров";
            this.ContourSchemeBtn.Theme = MetroFramework.MetroThemeStyle.Dark;
            this.ContourSchemeBtn.Click += new System.EventHandler(this.ContourSchemeBtn_Click);
            // 
            // MarkBtn
            // 
            this.MarkBtn.Location = new System.Drawing.Point(155, 165);
            this.MarkBtn.Name = "MarkBtn";
            this.MarkBtn.Size = new System.Drawing.Size(155, 23);
            this.MarkBtn.Style = MetroFramework.MetroColorStyle.Yellow;
            this.MarkBtn.TabIndex = 2;
            this.MarkBtn.Text = "Маркировка";
            this.MarkBtn.Theme = MetroFramework.MetroThemeStyle.Dark;
            this.MarkBtn.Click += new System.EventHandler(this.MarkBtn_Click);
            // 
            // ControlSchemeBtn
            // 
            this.ControlSchemeBtn.Location = new System.Drawing.Point(155, 194);
            this.ControlSchemeBtn.Name = "ControlSchemeBtn";
            this.ControlSchemeBtn.Size = new System.Drawing.Size(155, 23);
            this.ControlSchemeBtn.Style = MetroFramework.MetroColorStyle.Yellow;
            this.ControlSchemeBtn.TabIndex = 3;
            this.ControlSchemeBtn.Text = "Схема управления";
            this.ControlSchemeBtn.Theme = MetroFramework.MetroThemeStyle.Dark;
            this.ControlSchemeBtn.Click += new System.EventHandler(this.ControlSchemeBtn_Click);
            // 
            // TerminalsBtn
            // 
            this.TerminalsBtn.Location = new System.Drawing.Point(155, 223);
            this.TerminalsBtn.Name = "TerminalsBtn";
            this.TerminalsBtn.Size = new System.Drawing.Size(155, 23);
            this.TerminalsBtn.Style = MetroFramework.MetroColorStyle.Yellow;
            this.TerminalsBtn.TabIndex = 4;
            this.TerminalsBtn.Text = "Веса клемм";
            this.TerminalsBtn.Theme = MetroFramework.MetroThemeStyle.Dark;
            this.TerminalsBtn.Click += new System.EventHandler(this.TerminalsBtn_Click);
            // 
            // ChangeArticlesBtn
            // 
            this.ChangeArticlesBtn.Location = new System.Drawing.Point(155, 339);
            this.ChangeArticlesBtn.Name = "ChangeArticlesBtn";
            this.ChangeArticlesBtn.Size = new System.Drawing.Size(155, 23);
            this.ChangeArticlesBtn.Style = MetroFramework.MetroColorStyle.Yellow;
            this.ChangeArticlesBtn.TabIndex = 5;
            this.ChangeArticlesBtn.Text = "ChangeArticles";
            this.ChangeArticlesBtn.Theme = MetroFramework.MetroThemeStyle.Dark;
            this.ChangeArticlesBtn.Visible = false;
            this.ChangeArticlesBtn.Click += new System.EventHandler(this.ChangeArticlesBtn_Click);
            // 
            // SLineSchemeBtn
            // 
            this.SLineSchemeBtn.Location = new System.Drawing.Point(155, 252);
            this.SLineSchemeBtn.Name = "SLineSchemeBtn";
            this.SLineSchemeBtn.Size = new System.Drawing.Size(155, 23);
            this.SLineSchemeBtn.Style = MetroFramework.MetroColorStyle.Yellow;
            this.SLineSchemeBtn.TabIndex = 6;
            this.SLineSchemeBtn.Text = "Схема однолинейная";
            this.SLineSchemeBtn.Theme = MetroFramework.MetroThemeStyle.Dark;
            this.SLineSchemeBtn.Click += new System.EventHandler(this.SLineSchemeBtn_Click);
            // 
            // ChangeAttributesBtn
            // 
            this.ChangeAttributesBtn.Location = new System.Drawing.Point(155, 368);
            this.ChangeAttributesBtn.Name = "ChangeAttributesBtn";
            this.ChangeAttributesBtn.Size = new System.Drawing.Size(155, 23);
            this.ChangeAttributesBtn.Style = MetroFramework.MetroColorStyle.Yellow;
            this.ChangeAttributesBtn.TabIndex = 7;
            this.ChangeAttributesBtn.Text = "ChangeAttributes";
            this.ChangeAttributesBtn.Theme = MetroFramework.MetroThemeStyle.Dark;
            this.ChangeAttributesBtn.Visible = false;
            this.ChangeAttributesBtn.Click += new System.EventHandler(this.ChangeAttributesBtn_Click);
            // 
            // ChangeTextHeghtBtn
            // 
            this.ChangeTextHeghtBtn.Location = new System.Drawing.Point(137, 281);
            this.ChangeTextHeghtBtn.Name = "ChangeTextHeghtBtn";
            this.ChangeTextHeghtBtn.Size = new System.Drawing.Size(186, 23);
            this.ChangeTextHeghtBtn.Style = MetroFramework.MetroColorStyle.Yellow;
            this.ChangeTextHeghtBtn.TabIndex = 8;
            this.ChangeTextHeghtBtn.Text = "Изменить высоту и стиль текста";
            this.ChangeTextHeghtBtn.Theme = MetroFramework.MetroThemeStyle.Dark;
            this.ChangeTextHeghtBtn.Click += new System.EventHandler(this.ChangeTextHeghtBtn_Click);
            // 
            // MarkExportBtn
            // 
            this.MarkExportBtn.Location = new System.Drawing.Point(86, 310);
            this.MarkExportBtn.Name = "MarkExportBtn";
            this.MarkExportBtn.Size = new System.Drawing.Size(284, 23);
            this.MarkExportBtn.Style = MetroFramework.MetroColorStyle.Yellow;
            this.MarkExportBtn.TabIndex = 9;
            this.MarkExportBtn.Text = "Марировка со схемы управления (от 07.2017)";
            this.MarkExportBtn.Theme = MetroFramework.MetroThemeStyle.Dark;
            this.MarkExportBtn.Click += new System.EventHandler(this.MarkExportBtn_Click);
            // 
            // HelpBtn
            // 
            this.HelpBtn.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.HelpBtn.Location = new System.Drawing.Point(385, 28);
            this.HelpBtn.Name = "HelpBtn";
            this.HelpBtn.Size = new System.Drawing.Size(75, 23);
            this.HelpBtn.Style = MetroFramework.MetroColorStyle.Yellow;
            this.HelpBtn.TabIndex = 10;
            this.HelpBtn.Text = "Справка";
            this.HelpBtn.Theme = MetroFramework.MetroThemeStyle.Dark;
            this.HelpBtn.Click += new System.EventHandler(this.HelpBtn_Click);
            // 
            // ControlSchemeToTxtBtn
            // 
            this.ControlSchemeToTxtBtn.Location = new System.Drawing.Point(86, 136);
            this.ControlSchemeToTxtBtn.Name = "ControlSchemeToTxtBtn";
            this.ControlSchemeToTxtBtn.Size = new System.Drawing.Size(155, 23);
            this.ControlSchemeToTxtBtn.Style = MetroFramework.MetroColorStyle.Yellow;
            this.ControlSchemeToTxtBtn.TabIndex = 2;
            this.ControlSchemeToTxtBtn.Text = "Парсер схемы управления";
            this.ControlSchemeToTxtBtn.Theme = MetroFramework.MetroThemeStyle.Dark;
            this.ControlSchemeToTxtBtn.Click += new System.EventHandler(this.ControlSchemeToTxtBtn_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(460, 349);
            this.ControlBox = false;
            this.Controls.Add(this.HelpBtn);
            this.Controls.Add(this.MarkExportBtn);
            this.Controls.Add(this.ChangeTextHeghtBtn);
            this.Controls.Add(this.ChangeAttributesBtn);
            this.Controls.Add(this.SLineSchemeBtn);
            this.Controls.Add(this.ChangeArticlesBtn);
            this.Controls.Add(this.TerminalsBtn);
            this.Controls.Add(this.ControlSchemeBtn);
            this.Controls.Add(this.ControlSchemeToTxtBtn);
            this.Controls.Add(this.MarkBtn);
            this.Controls.Add(this.ContourSchemeBtn);
            this.Controls.Add(this.ConnectionSchemeBtn);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "MainForm";
            this.Resizable = false;
            this.Style = MetroFramework.MetroColorStyle.Yellow;
            this.Text = "AcElectricalSchemePlugin";
            this.Theme = MetroFramework.MetroThemeStyle.Dark;
            this.ResumeLayout(false);

        }

        #endregion

        private MetroFramework.Controls.MetroButton ConnectionSchemeBtn;
        private MetroFramework.Controls.MetroButton ContourSchemeBtn;
        private MetroFramework.Controls.MetroButton MarkBtn;
        private MetroFramework.Controls.MetroButton ControlSchemeBtn;
        private MetroFramework.Controls.MetroButton TerminalsBtn;
        private MetroFramework.Controls.MetroButton ChangeArticlesBtn;
        private MetroFramework.Controls.MetroButton SLineSchemeBtn;
        private MetroFramework.Controls.MetroButton ChangeAttributesBtn;
        private MetroFramework.Controls.MetroButton ChangeTextHeghtBtn;
        private MetroFramework.Controls.MetroButton MarkExportBtn;
        private MetroFramework.Controls.MetroButton HelpBtn;
        private MetroFramework.Controls.MetroButton ControlSchemeToTxtBtn;
    }
}
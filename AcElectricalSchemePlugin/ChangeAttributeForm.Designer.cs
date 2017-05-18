namespace AcElectricalSchemePlugin
{
    partial class ChangeAttributeForm
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
            this.attrListBox = new System.Windows.Forms.ListBox();
            this.attrNameListBox = new System.Windows.Forms.TextBox();
            this.okBtn = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // attrListBox
            // 
            this.attrListBox.FormattingEnabled = true;
            this.attrListBox.Location = new System.Drawing.Point(13, 13);
            this.attrListBox.Name = "attrListBox";
            this.attrListBox.Size = new System.Drawing.Size(216, 238);
            this.attrListBox.TabIndex = 0;
            // 
            // attrNameListBox
            // 
            this.attrNameListBox.Location = new System.Drawing.Point(261, 105);
            this.attrNameListBox.Name = "attrNameListBox";
            this.attrNameListBox.Size = new System.Drawing.Size(155, 20);
            this.attrNameListBox.TabIndex = 1;
            // 
            // okBtn
            // 
            this.okBtn.Location = new System.Drawing.Point(302, 131);
            this.okBtn.Name = "okBtn";
            this.okBtn.Size = new System.Drawing.Size(75, 23);
            this.okBtn.TabIndex = 2;
            this.okBtn.Text = "Изменить";
            this.okBtn.UseVisualStyleBackColor = true;
            this.okBtn.Click += new System.EventHandler(this.okBtn_Click);
            // 
            // ChangeAttributeForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(428, 261);
            this.Controls.Add(this.okBtn);
            this.Controls.Add(this.attrNameListBox);
            this.Controls.Add(this.attrListBox);
            this.Name = "ChangeAttributeForm";
            this.Text = "Изменение значение аттрибута";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox attrNameListBox;
        private System.Windows.Forms.Button okBtn;
        public System.Windows.Forms.ListBox attrListBox;
    }
}
namespace AcElectricalSchemePlugin
{
    partial class ChangeTextHeightForm
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
            this.txtBluePrintChBox = new System.Windows.Forms.CheckBox();
            this.mtxtBluePrintChBox = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.blocksLBox = new System.Windows.Forms.ListBox();
            this.selBlocksLBox = new System.Windows.Forms.ListBox();
            this.label2 = new System.Windows.Forms.Label();
            this.addBtn = new System.Windows.Forms.Button();
            this.deleteBtn = new System.Windows.Forms.Button();
            this.addAllBtn = new System.Windows.Forms.Button();
            this.deleteAllBtn = new System.Windows.Forms.Button();
            this.accBtn = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.mtxtBlocksChBox = new System.Windows.Forms.CheckBox();
            this.txtBlocksChBox = new System.Windows.Forms.CheckBox();
            this.label4 = new System.Windows.Forms.Label();
            this.compressionCoeffTBox = new System.Windows.Forms.TextBox();
            this.txtHeightTBox = new System.Windows.Forms.TextBox();
            this.tableChBox = new System.Windows.Forms.CheckBox();
            this.label5 = new System.Windows.Forms.Label();
            this.textStyleLBox = new System.Windows.Forms.ListBox();
            this.SuspendLayout();
            // 
            // txtBluePrintChBox
            // 
            this.txtBluePrintChBox.AutoSize = true;
            this.txtBluePrintChBox.Checked = true;
            this.txtBluePrintChBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.txtBluePrintChBox.Location = new System.Drawing.Point(202, 12);
            this.txtBluePrintChBox.Name = "txtBluePrintChBox";
            this.txtBluePrintChBox.Size = new System.Drawing.Size(201, 17);
            this.txtBluePrintChBox.TabIndex = 0;
            this.txtBluePrintChBox.Text = "Текст в чертеже (не считая блоки)";
            this.txtBluePrintChBox.UseVisualStyleBackColor = true;
            // 
            // mtxtBluePrintChBox
            // 
            this.mtxtBluePrintChBox.AutoSize = true;
            this.mtxtBluePrintChBox.Checked = true;
            this.mtxtBluePrintChBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.mtxtBluePrintChBox.Location = new System.Drawing.Point(202, 36);
            this.mtxtBluePrintChBox.Name = "mtxtBluePrintChBox";
            this.mtxtBluePrintChBox.Size = new System.Drawing.Size(210, 17);
            this.mtxtBluePrintChBox.TabIndex = 1;
            this.mtxtBluePrintChBox.Text = "МТекст в чертеже (не считая блоки)";
            this.mtxtBluePrintChBox.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 67);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(92, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Блоки в чертеже";
            // 
            // blocksLBox
            // 
            this.blocksLBox.FormattingEnabled = true;
            this.blocksLBox.Location = new System.Drawing.Point(15, 84);
            this.blocksLBox.Name = "blocksLBox";
            this.blocksLBox.Size = new System.Drawing.Size(249, 147);
            this.blocksLBox.TabIndex = 3;
            // 
            // selBlocksLBox
            // 
            this.selBlocksLBox.FormattingEnabled = true;
            this.selBlocksLBox.Location = new System.Drawing.Point(350, 84);
            this.selBlocksLBox.Name = "selBlocksLBox";
            this.selBlocksLBox.Size = new System.Drawing.Size(249, 147);
            this.selBlocksLBox.TabIndex = 5;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(347, 67);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(118, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Блоки для изменения";
            // 
            // addBtn
            // 
            this.addBtn.Location = new System.Drawing.Point(269, 102);
            this.addBtn.Name = "addBtn";
            this.addBtn.Size = new System.Drawing.Size(75, 23);
            this.addBtn.TabIndex = 6;
            this.addBtn.Text = "->";
            this.addBtn.UseVisualStyleBackColor = true;
            this.addBtn.Click += new System.EventHandler(this.addBtn_Click);
            // 
            // deleteBtn
            // 
            this.deleteBtn.Location = new System.Drawing.Point(269, 131);
            this.deleteBtn.Name = "deleteBtn";
            this.deleteBtn.Size = new System.Drawing.Size(75, 23);
            this.deleteBtn.TabIndex = 7;
            this.deleteBtn.Text = "<-";
            this.deleteBtn.UseVisualStyleBackColor = true;
            this.deleteBtn.Click += new System.EventHandler(this.deleteBtn_Click);
            // 
            // addAllBtn
            // 
            this.addAllBtn.Location = new System.Drawing.Point(269, 160);
            this.addAllBtn.Name = "addAllBtn";
            this.addAllBtn.Size = new System.Drawing.Size(75, 23);
            this.addAllBtn.TabIndex = 8;
            this.addAllBtn.Text = "-->>";
            this.addAllBtn.UseVisualStyleBackColor = true;
            this.addAllBtn.Click += new System.EventHandler(this.addAllBtn_Click);
            // 
            // deleteAllBtn
            // 
            this.deleteAllBtn.Location = new System.Drawing.Point(269, 189);
            this.deleteAllBtn.Name = "deleteAllBtn";
            this.deleteAllBtn.Size = new System.Drawing.Size(75, 23);
            this.deleteAllBtn.TabIndex = 8;
            this.deleteAllBtn.Text = "<<--";
            this.deleteAllBtn.UseVisualStyleBackColor = true;
            this.deleteAllBtn.Click += new System.EventHandler(this.deleteAllBtn_Click);
            // 
            // accBtn
            // 
            this.accBtn.Location = new System.Drawing.Point(404, 372);
            this.accBtn.Name = "accBtn";
            this.accBtn.Size = new System.Drawing.Size(110, 40);
            this.accBtn.TabIndex = 9;
            this.accBtn.Text = "Применить";
            this.accBtn.UseVisualStyleBackColor = true;
            this.accBtn.Click += new System.EventHandler(this.accBtn_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(339, 312);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(83, 13);
            this.label3.TabIndex = 11;
            this.label3.Text = "Размер текста";
            // 
            // mtxtBlocksChBox
            // 
            this.mtxtBlocksChBox.AutoSize = true;
            this.mtxtBlocksChBox.Checked = true;
            this.mtxtBlocksChBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.mtxtBlocksChBox.Location = new System.Drawing.Point(411, 261);
            this.mtxtBlocksChBox.Name = "mtxtBlocksChBox";
            this.mtxtBlocksChBox.Size = new System.Drawing.Size(112, 17);
            this.mtxtBlocksChBox.TabIndex = 13;
            this.mtxtBlocksChBox.Text = "МТекст в блоках";
            this.mtxtBlocksChBox.UseVisualStyleBackColor = true;
            // 
            // txtBlocksChBox
            // 
            this.txtBlocksChBox.AutoSize = true;
            this.txtBlocksChBox.Checked = true;
            this.txtBlocksChBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.txtBlocksChBox.Location = new System.Drawing.Point(411, 237);
            this.txtBlocksChBox.Name = "txtBlocksChBox";
            this.txtBlocksChBox.Size = new System.Drawing.Size(103, 17);
            this.txtBlocksChBox.TabIndex = 12;
            this.txtBlocksChBox.Text = "Текст в блоках";
            this.txtBlocksChBox.UseVisualStyleBackColor = true;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(467, 312);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(111, 13);
            this.label4.TabIndex = 15;
            this.label4.Text = "Коэффицент сжатия";
            // 
            // compressionCoeffTBox
            // 
            this.compressionCoeffTBox.Location = new System.Drawing.Point(470, 328);
            this.compressionCoeffTBox.Name = "compressionCoeffTBox";
            this.compressionCoeffTBox.Size = new System.Drawing.Size(110, 20);
            this.compressionCoeffTBox.TabIndex = 14;
            this.compressionCoeffTBox.Text = "1.0";
            this.compressionCoeffTBox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.compressionCoeffTBox_KeyPress);
            // 
            // txtHeightTBox
            // 
            this.txtHeightTBox.Location = new System.Drawing.Point(342, 328);
            this.txtHeightTBox.Name = "txtHeightTBox";
            this.txtHeightTBox.Size = new System.Drawing.Size(110, 20);
            this.txtHeightTBox.TabIndex = 10;
            this.txtHeightTBox.Text = "2.5";
            this.txtHeightTBox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtHeightTBox_KeyPress);
            // 
            // tableChBox
            // 
            this.tableChBox.AutoSize = true;
            this.tableChBox.Checked = true;
            this.tableChBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.tableChBox.Location = new System.Drawing.Point(411, 290);
            this.tableChBox.Name = "tableChBox";
            this.tableChBox.Size = new System.Drawing.Size(71, 17);
            this.tableChBox.TabIndex = 16;
            this.tableChBox.Text = "Таблицы";
            this.tableChBox.UseVisualStyleBackColor = true;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(12, 243);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(74, 13);
            this.label5.TabIndex = 2;
            this.label5.Text = "Стили текста";
            // 
            // textStyleLBox
            // 
            this.textStyleLBox.FormattingEnabled = true;
            this.textStyleLBox.Location = new System.Drawing.Point(15, 260);
            this.textStyleLBox.Name = "textStyleLBox";
            this.textStyleLBox.Size = new System.Drawing.Size(144, 147);
            this.textStyleLBox.TabIndex = 3;
            // 
            // ChangeTextHeightForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(615, 419);
            this.Controls.Add(this.tableChBox);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.compressionCoeffTBox);
            this.Controls.Add(this.mtxtBlocksChBox);
            this.Controls.Add(this.txtBlocksChBox);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtHeightTBox);
            this.Controls.Add(this.accBtn);
            this.Controls.Add(this.deleteAllBtn);
            this.Controls.Add(this.addAllBtn);
            this.Controls.Add(this.deleteBtn);
            this.Controls.Add(this.addBtn);
            this.Controls.Add(this.selBlocksLBox);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.textStyleLBox);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.blocksLBox);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.mtxtBluePrintChBox);
            this.Controls.Add(this.txtBluePrintChBox);
            this.Name = "ChangeTextHeightForm";
            this.Text = "Изменение размера текста";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckBox txtBluePrintChBox;
        private System.Windows.Forms.CheckBox mtxtBluePrintChBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListBox blocksLBox;
        private System.Windows.Forms.ListBox selBlocksLBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button addBtn;
        private System.Windows.Forms.Button deleteBtn;
        private System.Windows.Forms.Button addAllBtn;
        private System.Windows.Forms.Button deleteAllBtn;
        private System.Windows.Forms.Button accBtn;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.CheckBox mtxtBlocksChBox;
        private System.Windows.Forms.CheckBox txtBlocksChBox;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox compressionCoeffTBox;
        private System.Windows.Forms.TextBox txtHeightTBox;
        private System.Windows.Forms.CheckBox tableChBox;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ListBox textStyleLBox;
    }
}
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;

namespace AcElectricalSchemePlugin
{
    public partial class ChangeTextHeightForm : Form
    {

        public ChangeTextHeightForm(List<string> blocks, List<string> styles)
        {
            InitializeComponent();
            blocksLBox.Items.AddRange(blocks.ToArray());
            textStyleLBox.Items.AddRange(styles.ToArray());
            if (textStyleLBox.Items.Count>0)
            {
                int index = textStyleLBox.Items.IndexOf("GOSTA-2.5-1");
                if (index >= 0) textStyleLBox.SelectedIndex = index;
                else textStyleLBox.SelectedIndex = 0;
            }
        }

        private void addBtn_Click(object sender, EventArgs e)
        {
            if (blocksLBox.SelectedIndex>=0)
            {
                int oldIndex = blocksLBox.SelectedIndex;
                selBlocksLBox.Items.Add(blocksLBox.SelectedItem);
                blocksLBox.Items.RemoveAt(blocksLBox.SelectedIndex);
                if (blocksLBox.Items.Count > 0)
                    if (oldIndex - 1 == blocksLBox.Items.Count - 1) blocksLBox.SelectedIndex = oldIndex - 1;
                    else blocksLBox.SelectedIndex = oldIndex;
            }
        }

        private void deleteBtn_Click(object sender, EventArgs e)
        {
            if (selBlocksLBox.SelectedIndex >= 0)
            {
                int oldIndex = selBlocksLBox.SelectedIndex;
                blocksLBox.Items.Add(selBlocksLBox.SelectedItem);
                selBlocksLBox.Items.RemoveAt(selBlocksLBox.SelectedIndex);
                if (selBlocksLBox.Items.Count > 0)
                    if (oldIndex - 1 == selBlocksLBox.Items.Count - 1) selBlocksLBox.SelectedIndex = oldIndex - 1;
                    else selBlocksLBox.SelectedIndex = oldIndex;
            }
        }

        private void addAllBtn_Click(object sender, EventArgs e)
        {
            if (blocksLBox.Items.Count >= 1)
            {
                selBlocksLBox.Items.AddRange(blocksLBox.Items);
                blocksLBox.Items.Clear();
            }
        }

        private void deleteAllBtn_Click(object sender, EventArgs e)
        {
            if (selBlocksLBox.Items.Count >= 1)
            {
                blocksLBox.Items.AddRange(selBlocksLBox.Items);
                selBlocksLBox.Items.Clear();
            }
        }

        private void accBtn_Click(object sender, EventArgs e)
        {
            if (txtHeightTBox!=null)
            {
                double txtHeight = 0;
                double.TryParse(txtHeightTBox.Text.Replace(",", CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator), out txtHeight);
                double compCoeff = 0;
                double.TryParse(compressionCoeffTBox.Text.Replace(",", CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator), out compCoeff);
                if (txtHeight != 0)
                {
                    if (compCoeff != 0)
                    {
                        if (textStyleLBox.SelectedIndex >= 0)
                        {
                            List<string> namesOfBlocks = new List<string>();
                            for (int i = 0; i < selBlocksLBox.Items.Count; i++)
                            {
                                namesOfBlocks.Add(selBlocksLBox.Items[i].ToString());
                            }
                            TextHeightChangeClass.ChangeTextHeight(mtxtBluePrintChBox.Checked, txtBluePrintChBox.Checked, mtxtBlocksChBox.Checked, txtBlocksChBox.Checked, tableChBox.Checked, txtHeight, compCoeff, namesOfBlocks, textStyleLBox.SelectedItem.ToString());
                            this.Close();
                        }
                    }
                    else MessageBox.Show("Введён некорректный коэффцицент сжатия");
                }
                else MessageBox.Show("Введён некорректный размер текста");
            }
        }

        private void txtHeightTBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsNumber(e.KeyChar) || e.KeyChar == '.' || e.KeyChar == ',')
                e.Handled = false;
            else
                e.Handled = true;
        }

        private void compressionCoeffTBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsNumber(e.KeyChar) || e.KeyChar=='.' || e.KeyChar==',')
                e.Handled = false;
            else
                e.Handled = true;
        }
    }
}

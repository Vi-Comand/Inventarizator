using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Инвентаризатор
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {


            if (checkBox3.Checked)
            {
                label3.Visible = true;
                textBox1.Visible = true;
            }
            else
            {
                label3.Visible = false;
                textBox1.Visible = false;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            // получаем выбранный файл
            string filename = openFileDialog1.FileName;
            // сохраняем текст в файл
            label1.Text = filename;
            if (label1.Text != "" && label2.Text != "")
            {
                groupBox1.Visible = true;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            // получаем выбранный файл
            string filename = openFileDialog1.FileName;
            // сохраняем текст в файл
            label2.Text = filename;

            if (label1.Text != "" && label2.Text != "")
            {
                groupBox1.Visible = true;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            groupBox3.Visible = true;
        }
    }
}

using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
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

        public class ExcelData
        {
            public string Id { get; set; }
            public string Name { get; set; }
            public string Gender { get; set; }
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
            label4.Text = DateTime.Now.ToString();
            //Table_To_Object_Test();
            label4.Text += DateTime.Now.ToString();
        }

        public void Table_To_Object_Test()
        {
            //Create a test file
            var fi = new FileInfo(label1.Text);

            using (ExcelPackage package = new ExcelPackage(fi))
            {
                var workbook = package.Workbook;
                var worksheet = workbook.Worksheets.First();
                int colCount = 6;  //get Column Count
                int rowCount = worksheet.Dimension.End.Row;
                //var ThatList = worksheet.Tables.First().ConvertTableToObjects<ExcelData>();
                List<ExcelData> ex = new List<ExcelData>();
                string q;
                string qq;
                string qqq;
                for (int row = 1; row <= rowCount; row++)
                {
                    try
                    {
                        q = worksheet.Cells[row, 3].Value.ToString();
                    }
                    catch
                    {
                        q = "";
                    }
                    try
                    {
                        qq = worksheet.Cells[row, 1].Value.ToString();
                    }
                    catch
                    {
                        qq = "";
                    }
                    try
                    {
                        qqq = worksheet.Cells[row, 6].Value.ToString();
                    }
                    catch
                    {
                        qqq = "";
                    }
                    ex.Add(new ExcelData()
                    {
                        Id = q,
                        Name = qq,
                        Gender = qqq
                    });
                }
                // MessageBox.Show(ex.Count.ToString());
                //package.Save();
            }
        }
    }
}

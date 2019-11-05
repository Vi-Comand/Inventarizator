using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System.IO;

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
          
            FormirList List = new FormirList();
            List<FormirList> list1 = List.FormirListExcel(label1.Text);
            List<FormirList> list2 = List.FormirListExcel(label2.Text);
            Lists op1 = new Lists();

            op1.list1 = list1.Where(x=>x.id1!=0).ToList();
            op1.list2 = list2.Where(x => x.id1 != 0).ToList();
            Lists op= NaFullSovpadenie(list1, list2);
          op= NaOdinnakovoeStrihkodPolka(op.list1,op.list2);

            op1.doc1 = op.doc1;
            op1.doc2 = op.doc2;
            IzmKolvoVFiles(op1);
           
            }
        private void Excel(string path, List<FormirList> list)
        {
            
            //FileInfo newFile = null;
            /*if (!File.Exists(path + "\\testsheet2.xlsx"))
            newFile = new FileInfo(path + "\\testsheet2.xlsx");
            else
                return newFile;*/
            using (ExcelPackage package = new ExcelPackage(new FileInfo(path)))
            {

                ExcelWorksheet ws = package.Workbook.Worksheets.Add("");
                int i = 1;
                foreach (var row in list)
                { i++;

                    ws.Cells[i,1].Value =row.strihKod ;

                    ws.Cells[i, 1].Style.Font.Color.SetColor(Color.Red);
                    
                }

                package.Save();
            }

        }


        private void NaOdinnakovoeStrihkodKolvo(List<FormirList> list1, List<FormirList> list2)
        {

            var leftList =
(from l1 in list1
 join l2 in list2 on l1.strihKod +"/" + l1.kol_vo equals l2.strihKod + "/" + l2.kol_vo
 select new FormirList
 {id1=l1.id1,
     strihKod = (l1 != null ? l1.strihKod : l2.strihKod),
     shkaf = l1.shkaf,
     polka = l1.polka,
     kol_vo = l1.kol_vo,
     id2= (l2 != null ? l2.id1 : 0),
     shkaf1 = (l2 != null ? l2.shkaf : null),
     polka1 = (l2 != null ? l2.polka : null),
     kol_vo1 = (l2 != null ? l2.kol_vo : 0)
 }


);
            //int[] mas1 = new int[leftList.Count()];
            //int[] mas2 = new int[leftList.Count()];
            //int i = 0;
            //foreach (var row in leftList)
            //{
            //    mas1[i] = row.l1;
            //    mas2[i] = row.l2;
            //    i++;
            //}

            //list1.RemoveAll(x => mas1.Contains(x.id1));
            //list2.RemoveAll(x => mas2.Contains(x.id1));

        }
        private Lists NaOdinnakovoeStrihkodPolka(List<FormirList> list1, List<FormirList> list2)
        {

            var leftList =
(from l1 in list1
 join l2 in list2 on l1.strihKod + "/" + l1.polka equals l2.strihKod + "/" + l2.polka
 select new FormirList
 {
     id1 = l1.id1,
     strihKod = (l1 != null ? l1.strihKod : l2.strihKod),
     shkaf = l1.shkaf,
     polka = l1.polka,
     kol_vo = l1.kol_vo,
     id2 = (l2 != null ? l2.id1 : 0),
     shkaf1 = (l2 != null ? l2.shkaf : null),
     polka1 = (l2 != null ? l2.polka : null),
     kol_vo1 = (l2 != null ? l2.kol_vo : 0)
 }).ToList();



            int[] mas1 = new int[leftList.Count()];
            int[] mas2 = new int[leftList.Count()];
            int i = 0;

            List<IdZnach> doc1 = new List<IdZnach>();
            List<IdZnach> doc2 = new List<IdZnach>();
            //List<FormirList> leftList = leftLis.Where(x => x.id1 != 0).ToList();
           for(int j=0;j<leftList.Count();j++)
            {
                if (leftList[j].kol_vo == leftList[j].kol_vo1)
                {
                    mas1[i] = leftList[j].id1;
                    mas2[i] = leftList[j].id2;
                    i++;
                    leftList.Remove(leftList[j]);
                }
                else
                {
                    int ch1 = 0;
                    int ch2 = 0;

                    if (leftList[j].kol_vo > leftList[j].kol_vo1)
                    {
                        ch1 = leftList[j].kol_vo;
                        ch2 = leftList[j].kol_vo1;
                    }
                    else
                    {
                        ch1 = leftList[j].kol_vo1;
                        ch2 = leftList[j].kol_vo;
                    }
                    if (ch1 > 0 && ch2 > 0)
                    {
                        double per = ch2 / 100;
                        per = (ch1 - ch2) / per;

                        if (per < Convert.ToDouble(textBox1.Text))
                        {

                            per = (ch1 + ch2) / 2;
                            int konec = (int)Math.Round(per, 0);
                            doc1.Add(new IdZnach(leftList[j].id1, konec));
                            doc2.Add(new IdZnach(leftList[j].id2, konec));
                            leftList.Remove(leftList[j]);
                            mas1[i] = leftList[j].id1;
                            mas2[i] = leftList[j].id2;
                            i++;
                        }
                    }

                

                }
            }
            list1.RemoveAll(x => mas1.Contains(x.id1));
            list2.RemoveAll(x => mas2.Contains(x.id1));
            Lists op = new Lists();
            op.list1 = list1;
            op.list2 = list2;
            op.doc1 = doc1;
            op.doc2 = doc2;


            return op;
        }
        private void IzmKolvoVFiles(Lists op)
        {
            foreach (var str in op.doc1)
            {
                var a = op.list1.Where(x => x.id1 == str.id).First();
                a.kol_vo = str.znach;
            }
            foreach (var str in op.doc2)
            {
                var a = op.list2.Where(x => x.id1 == str.id).First();
                a.kol_vo = str.znach;
            }

        }
        private Lists NaFullSovpadenie(List<FormirList> list1, List<FormirList> list2)
        {

            var leftList =
            (from l1 in list1
              join l2 in list2 on l1.strihKod +l1.shkaf+ l1.polka + l1.kol_vo equals l2.strihKod + l2.shkaf + l2.polka + l2.kol_vo
                select new IdList
 {
     l1= l1.id1,
     l2=l2.id1
 }


);
            int[] mas1=new int[leftList.Count()];
            int[] mas2= new int[leftList.Count()]; 
            int i = 0;
            foreach (var row in leftList)
            {
                mas1[i] = row.l1;
                 mas2[i] = row.l2;
                i++;
            }

            Lists op = new Lists();

            list1.RemoveAll(x => mas1.Contains(x.id1));
            list2.RemoveAll(x => mas2.Contains(x.id1));
            op.list1 = list1;
            op.list2 = list2;
            /* var leftList =
       (from l1 in list1
     join l2 in list2 on l1.strihKod+l1.polka+l1.kol_vo equals l2.strihKod+l2.polka+l2.kol_vo 

        select new FormirList
        {
            strihKod = (l1 != null ? l1.strihKod : l2.strihKod),
            shkaf = l1.shkaf,
           polka = l1.polka,
            kol_vo = l1.kol_vo,

            shkaf1 = (l2!=null?l2.shkaf:null),
            polka1 = (l2 != null ? l2.polka : null),
            kol_vo1 = (l2 != null ? l2.kol_vo : null)
        });
             var  c = list2.FindAll(w => (list1.Find(x => w.strihKod == x.strihKod) == null));
               var c1 = list1.FindAll(w => (list2.Find(x => w.strihKod == x.strihKod) == null));
               List<FormirList> c2 = new List<FormirList>();
               c2.AddRange(c);
               c2.AddRange(c1);*/
            return op;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            /* FormirList List = new FormirList();
                 //  string[,] mas = List.mas(100000, 4);
                    List < FormirList > Lise=  List.FormirListExcel(@"C:\Users\hozja_d_e\Desktop\100.xlsx");*/
            double a = 3.556666;
            int C =(int) Math.Round(a,0);
        }
    }
}

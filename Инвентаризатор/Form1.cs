﻿using OfficeOpenXml;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
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
        public string path;
        private void button3_Click(object sender, EventArgs e)
        {
            groupBox3.Visible = true;
            path = Directory.GetCurrentDirectory() + "\\" + DateTime.Now.ToFileTime();
            DirectoryInfo dirInfo = new DirectoryInfo(path);
            dirInfo.Create();
            FormirList List = new FormirList();
            List<FormirList> list1 = List.FormirListExcel(label1.Text);
            List<FormirList> list2 = List.FormirListExcel(label2.Text);
            Lists op1 = new Lists();

            op1.list1 = list1.Where(x => x.id1 != 0).ToList();
            op1.list2 = list2.Where(x => x.id1 != 0).ToList();
            Lists op = NaFullSovpadenie(list1, list2);
            op = NaOdinnakovoeStrihkodPolka(op.list1, op.list2);

            op1.doc1 = op.doc1;
            op1.doc2 = op.doc2;
            IzmKolvoVFiles(op1);
            MessageBox.Show("Готово");
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
                {
                    i++;

                    ws.Cells[i, 1].Value = row.strihKod;

                    ws.Cells[i, 1].Style.Font.Color.SetColor(Color.Red);

                }

                package.Save();
            }

        }


        private void NaOdinnakovoeStrihkodKolvo(List<FormirList> list1, List<FormirList> list2)
        {

            var leftList =
(from l1 in list1
 join l2 in list2 on l1.strihKod + "/" + l1.kol_vo equals l2.strihKod + "/" + l2.kol_vo
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
     name = l1.name,

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
            for (int j = 0; j < leftList.Count(); j++)
            {
                if (leftList[j].kol_vo == leftList[j].kol_vo1)
                {
                    mas1[i] = leftList[j].id1;
                    mas2[i] = leftList[j].id2;
                 
                    leftList.Remove(leftList[j]);
                    i++;
                    j--;
                }
                else
                {
                    double ch1 = 0;
                    double ch2 = 0;

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

                            mas1[i] = leftList[j].id1;
                            mas2[i] = leftList[j].id2;
                            leftList.Remove(leftList[j]);
                            j--;
                            i++;
                        }
                        else
                        {
                            mas1[i] = leftList[j].id1;
                            mas2[i] = leftList[j].id2;
                            i++;
                        }

                    }



                }
            }
            list1.RemoveAll(x => mas1.Contains(x.id1));
            list2.RemoveAll(x => mas2.Contains(x.id1));


            var StrihKodKolvoList =
(from l1 in list1
 join l2 in list2 on l1.strihKod + "/" + l1.kol_vo equals l2.strihKod + "/" + l2.kol_vo
 select new FormirList
 {
     name = l1.name,

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

            mas1 = new int[StrihKodKolvoList.Count()];
            mas2 = new int[StrihKodKolvoList.Count()];

            for (int j = 0; j < StrihKodKolvoList.Count(); j++)
            {

                mas1[j] = StrihKodKolvoList[j].id1;
                mas2[j] = StrihKodKolvoList[j].id2;
            }


            list1.RemoveAll(x => mas1.Contains(x.id1));
            list2.RemoveAll(x => mas2.Contains(x.id1));




            var StrihKod =
(from l1 in list1
join l2 in list2 on l1.strihKod equals l2.strihKod
select new FormirList
{
name = l1.name,

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

            mas1 = new int[StrihKod.Count()];
            mas2 = new int[StrihKod.Count()];

            for (int j = 0; j < StrihKod.Count(); j++)
            {

                mas1[j] = StrihKod[j].id1;
                mas2[j] = StrihKod[j].id2;
            }


            list1.RemoveAll(x => mas1.Contains(x.id1));
            list2.RemoveAll(x => mas2.Contains(x.id1));




            Lists op = new Lists();






            op.doc1 = doc1;
            op.doc2 = doc2;
            using (var package = new ExcelPackage(new FileInfo(Directory.GetCurrentDirectory() + "\\1s.xlsx")))
            {
                int n_str = 2;
                var workSheet = package.Workbook.Worksheets[1];

                foreach (var str in list1)
                {
                    workSheet.Cells[n_str, 1].Value = str.id1;
                    workSheet.Cells[n_str, 2].Value = str.strihKod;
                    workSheet.Cells[n_str, 3].Value = str.shkaf;
                    workSheet.Cells[n_str, 4].Value = str.polka;
                    workSheet.Cells[n_str, 5].Value = str.kol_vo;
                    n_str++;

                }
                string name = path + "\\Присутствуют только в файле 1.xlsx";
                package.SaveAs(new FileInfo(name));
            }
            using (var package = new ExcelPackage(new FileInfo(Directory.GetCurrentDirectory() + "\\1s.xlsx")))
            {
                int n_str = 2;
                var workSheet = package.Workbook.Worksheets[1];

                foreach (var str in list2)
                {
                    workSheet.Cells[n_str, 1].Value = str.id1;
                    workSheet.Cells[n_str, 2].Value = str.strihKod;
                    workSheet.Cells[n_str, 3].Value = str.shkaf;
                    workSheet.Cells[n_str, 4].Value = str.polka;
                    workSheet.Cells[n_str, 5].Value = str.kol_vo;
                    n_str++;

                }
                string name = path + "\\Присутствуют только в файле 2.xlsx";
                package.SaveAs(new FileInfo(name));
            }




            using (var package = new ExcelPackage(new FileInfo(Directory.GetCurrentDirectory() + "\\2s.xlsx")))
            {
                int n_str = 2;
                var workSheet = package.Workbook.Worksheets[1];

                foreach (var str in leftList)
                {
                    workSheet.Cells[n_str, 1].Value = str.name;
                    workSheet.Cells[n_str, 2].Value = str.id1;
                    workSheet.Cells[n_str, 3].Value = str.id2;
                    workSheet.Cells[n_str, 4].Value = str.strihKod;
                    workSheet.Cells[n_str, 5].Value = str.shkaf;
                    workSheet.Cells[n_str, 6].Value = str.polka;
                    workSheet.Cells[n_str, 7].Value = str.polka1;
                    workSheet.Cells[n_str, 8].Value = str.kol_vo;
                    workSheet.Cells[n_str, 9].Value = str.kol_vo1;
                    workSheet.Cells[n_str, 8].Style.Font.Color.SetColor(Color.Red);

                    workSheet.Cells[n_str, 9].Style.Font.Color.SetColor(Color.Red);
                    n_str++;

                }
                foreach (var str in StrihKodKolvoList)
                {


                    workSheet.Cells[n_str, 1].Value = str.name;
                    workSheet.Cells[n_str, 2].Value = str.id1;
                    workSheet.Cells[n_str, 3].Value = str.id2;
                    workSheet.Cells[n_str, 4].Value = str.strihKod;
                    workSheet.Cells[n_str, 5].Value = str.shkaf;
                    workSheet.Cells[n_str, 6].Value = str.polka;
                    workSheet.Cells[n_str, 6].Style.Font.Color.SetColor(Color.Red);
                    workSheet.Cells[n_str, 7].Value = str.polka1;
                    workSheet.Cells[n_str, 7].Style.Font.Color.SetColor(Color.Red);
                    workSheet.Cells[n_str, 8].Value = str.kol_vo;
                    workSheet.Cells[n_str, 9].Value = str.kol_vo1;
                    n_str++;




                }
                foreach (var str in StrihKod)
                {


                    workSheet.Cells[n_str, 1].Value = str.name;
                    workSheet.Cells[n_str, 2].Value = str.id1;
                    workSheet.Cells[n_str, 3].Value = str.id2;
                    workSheet.Cells[n_str, 4].Value = str.strihKod;
                    workSheet.Cells[n_str, 5].Value = str.shkaf;
                    workSheet.Cells[n_str, 6].Value = str.polka;
                    workSheet.Cells[n_str, 6].Style.Font.Color.SetColor(Color.Red);
                    workSheet.Cells[n_str, 7].Value = str.polka1;
                    workSheet.Cells[n_str, 7].Style.Font.Color.SetColor(Color.Red);
                    workSheet.Cells[n_str, 8].Value = str.kol_vo;
                    workSheet.Cells[n_str, 9].Value = str.kol_vo1;
                    workSheet.Cells[n_str, 8].Style.Font.Color.SetColor(Color.Red);
                    workSheet.Cells[n_str, 9].Style.Font.Color.SetColor(Color.Red);
                    n_str++;




                }
                string name = path + "\\Файл расхождений по полкам и количеству.xlsx";
                package.SaveAs(new FileInfo(name));
               
            }
            label4.Text = "Количество расхождений по полкам: " + (StrihKodKolvoList.Count() + StrihKod.Count()) + "\n Расхождение по количеству: " + (leftList.Count() + StrihKod.Count());
            return op;
        }
        private void IzmKolvoVFiles(Lists op)
        {
            label4.Text = label4.Text + "\n Количиство корректировок в документах: " + op.doc1.Count();

            using (var package = new ExcelPackage(new FileInfo(label1.Text)))
            {
                var workSheet = package.Workbook.Worksheets[1];

                foreach (var str in op.doc1)
                {

                    workSheet.Cells[str.id, 7].Value = str.znach;



                }
                string name = path + "\\Файл коректировки количества 1.xlsx";
                package.SaveAs(new FileInfo(name));
            }
            using (var package = new ExcelPackage(new FileInfo(label2.Text)))
            {
                var workSheet = package.Workbook.Worksheets[1];

                foreach (var str in op.doc2)
                {

                    workSheet.Cells[str.id, 7].Value = str.znach;



                }
                string name = path + "\\Файл коректировки количества 2.xlsx";
                package.SaveAs(new FileInfo(name));
            }
        }

        private Lists NaFullSovpadenie(List<FormirList> list1, List<FormirList> list2)
        {

            var leftList =
            (from l1 in list1
             join l2 in list2 on l1.strihKod + l1.shkaf + l1.polka + l1.kol_vo equals l2.strihKod + l2.shkaf + l2.polka + l2.kol_vo
             select new IdList
             {
                 l1 = l1.id1,
                 l2 = l2.id1
             }
             

);

            int kol  = leftList.Count();


            while (leftList.Count() > 0)
            {
                kol = leftList.Count();

                List<IdList> tis = leftList.Take(100000).ToList();
                // int i = 0;

                int[] mas1 = new int[100000];
                int[] mas2 = new int[100000];
                int i = 0;
                while (tis.Count()>0)
                {
                    mas1[i] = tis[0].l1;
                    mas2[i] = tis[0].l2;
                    tis.Remove(tis[0]);
                    i++;
                }

                list1.RemoveAll(x => mas1.Contains(x.id1));
                list2.RemoveAll(x => mas2.Contains(x.id1));
            }
           
          

            Lists op = new Lists();

          //  list1.RemoveAll(x => mas1.Contains(x.id1));
          //  list2.RemoveAll(x => mas2.Contains(x.id1));
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
            int C = (int)Math.Round(a, 0);
        }
    }
}

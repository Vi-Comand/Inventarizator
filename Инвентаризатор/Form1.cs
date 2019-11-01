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

            NaFullSovpadenie(list1, list2);
       
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


            foreach (var str in leftList)
            {
                if (str.kol_vo == str.kol_vo1)
                {
                    mas1[i] = str.id1;
                    mas2[i] = str.id2;
                    i++;
                    leftList.Remove(str);
                }
                else
                {       int ch1 = 0;
                        int ch2 = 0;

                    if (str.kol_vo > str.kol_vo1)
                    {   ch1 = str.kol_vo;
                        ch2 = str.kol_vo1; }
                    else
                    {
                        ch1 = str.kol_vo1;
                        ch2 = str.kol_vo;
                    }
                    if (ch1 > 0 && ch2>0)
                    {
                        double per = ch2 / 100;
                        per = (ch1 - ch2) / per;

                        if (per<Convert.ToDouble(textBox1.Text))
                        {

                            per = (ch1 + ch2) / 2;
                            int  konec = (int)Math.Round(per, 0);
                           
                            leftList.Remove(str);
                            mas1[i] = str.id1;
                            mas2[i] = str.id2;
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



            return op;
        }
        private void NaFullSovpadenie(List<FormirList> list1, List<FormirList> list2)
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

            list1.RemoveAll(x => mas1.Contains(x.id1));
            list2.RemoveAll(x => mas2.Contains(x.id1));

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

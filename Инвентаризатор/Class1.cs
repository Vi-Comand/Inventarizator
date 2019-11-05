using LinqToExcel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Инвентаризатор
{
    class Lists
    {
        public List<FormirList> list1{ get; set; }
        public List<FormirList> list2 { get; set; }
        public List<IdZnach> doc1 { get; set; }
        public List<IdZnach> doc2 { get; set; }

    }


    class IdList
    {
        public int l1 { get; set; }
        public int l2 { get; set; }
    }
    class IdZnach
    {
        public IdZnach(int id,int znach)
        { this.id = id;
            this.znach = znach;
        }
        public int id { get; set; }
    public int znach { get; set; }
}

class FormirList
    {

        public int id1 { get; set; }
        public int id2 { get; set; }
        public string name { get; set; }
        public string strihKod { get; set; }
        public string articul { get; set; }
        public string kol_vof { get; set; }
        public string shkaf { get; set; }
        public string polka { get; set; }
        public int kol_vo { get; set; }
        public string vrem { get; set; }
        public string itog { get; set; }
        public string name1 { get; set; }
        public string shkaf1 { get; set; }
        public string polka1 { get; set; }
        public int kol_vo1 { get; set; }
        public string vrem1 { get; set; }
        public string articul1 { get; set; }
        public string kol_vof1 { get; set; }
        public string itog1{ get; set; }
        //public FormirList(string strihKod,
        //   string shkaf,
        //   string polka,
        //   string kol_vo)
        //  {
        //      this.strihKod = strihKod;
        //      this.shkaf = shkaf;
        //      this.polka = polka;
        //      this.kol_vo = kol_vo;
        //  }
        public string[,] mas(int str, int col)
        {
            Random rand = new Random();
            string[,] mas=new string[str,col];
            for (int i = 0; i < str; i++)
              
            for (int j = 0; j <col; j++)
                {
                    mas[i, j] = rand.Next(0, 999999).ToString();
                    }
            return mas;
        }


        public  List<FormirList> FormirListExcel(string path)
        {
          
            var excel = new ExcelQueryFactory(path);

            var info = (from l1 in excel.Worksheet(0)

                        select new FormirList

                        {


                            strihKod = l1[2].ToString(),
                            shkaf = l1[4].ToString(),
                            polka = l1[5].ToString(),
                            kol_vo = l1[6].ToString() != "" ? Convert.ToInt32(l1[6]):0
                        }).ToList();
            

               info = info.Where(x => x.strihKod != "").ToList();
            int i = 2;
            foreach (var str in info)
            {
                str.id1 = i++;
            }
                return info;
            }



            public List<FormirList> FormirListExcel(string[,] mas, int strok)
        {
            FormirList str;
            List<FormirList> listExcel = new List<FormirList>();
            for (int i = 0; i < strok; i++)
            {
                str = new FormirList();
                str.strihKod = mas[i, 0].ToString();
                str.shkaf = mas[i, 1].ToString();
                str.polka = mas[i, 2].ToString();
                str.kol_vo = Convert.ToInt32(mas[i, 3]);
                listExcel.Add(str);

            }
            return listExcel;

        }


        public List<FormirList> FormirListExcel(Excel.Worksheet sheet)
        {
            FormirList str ;


            List<FormirList> listExcel = new List<FormirList>();
            for(int i=2;i<sheet.UsedRange.Rows.Count; i++)
            {
                if (sheet.Cells[i, 3].Value != null)
                {
                    str = new FormirList();
                    str.strihKod = sheet.Cells[i, 3].Value.ToString();
                    str.shkaf = sheet.Cells[i, 5].Value.ToString();
                    str.polka = sheet.Cells[i, 6].Value.ToString();
                    str.kol_vo = sheet.Cells[i, 7].Value.ToString();
                    listExcel.Add(str);
                }
                else
                {
                    break;
                }

            }
            return listExcel; }

    }
}

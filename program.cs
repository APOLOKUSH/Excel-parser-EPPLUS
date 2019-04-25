using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.IO;


namespace final
{


    public class test
    {
        public static List<List<string>> excelData = new List<List<string>>();
    }

    class Program
    {

        static void Main(string[] args)
        {
            //List<List<string>> excelData = new List<List<string>>();



            first();
            second();
            third();






            //List<string> v1 = new List<string>();
            //List<string> v2 = new List<string>();
            //List<string> v3 = new List<string>();

            //FileInfo file = new FileInfo(@""); //File Path

            //ExcelPackage Package = new ExcelPackage(file);
            ////ExcelWorksheet worksheet = Package.Workbook.Worksheets.Add("Sheet 1");
            //int c = Package.Workbook.Worksheets.Count;
            //ExcelWorksheet worksheet = Package.Workbook.Worksheets[1];


               //
        }
       
        public static void first()
        {
            List<string> v1 = new List<string>();

            FileInfo file = new FileInfo(@"C"); //File Pah

            ExcelPackage Package = new ExcelPackage(file);
            //ExcelWorksheet worksheet = Package.Workbook.Worksheets.Add("Sheet 1");
            int c = Package.Workbook.Worksheets.Count;
            ExcelWorksheet worksheet = Package.Workbook.Worksheets[1];


            int j = 1;
            int i = 9;

            string s = worksheet.Cells["A1"].Value.ToString();
            //string a = worksheet.Cells["D1"].Value.ToString();

            //worksheet.Cells[i, j].ToString();
            while (true)
            {
                if (worksheet.Cells[i, j].Value == null)
                    break;
                else
                    v1.Add(worksheet.Cells[i, j].Value.ToString());
                i++;
            }

            test.excelData.Add(v1);

            
        }


        public static void second()
        {
            List<string> v2 = new List<string>();

            FileInfo file = new FileInfo(@""); //File Path

            ExcelPackage Package = new ExcelPackage(file);
            //ExcelWorksheet worksheet = Package.Workbook.Worksheets.Add("Sheet 1");
            int p = Package.Workbook.Worksheets.Count;
            ExcelWorksheet worksheet = Package.Workbook.Worksheets[1];

            int k = 9;//4
            int z = 4;

            string a = worksheet.Cells["D9"].Value.ToString();

            //worksheet.Cells[i, j].ToString();
            while (true)
            {
                if (worksheet.Cells[k, z].Value == null)
                    break;
                else
                    v2.Add(worksheet.Cells[k, z].Value.ToString());
                k++;
            }
            test.excelData.Add(v2);
        }

        public static void third()
        {
            List<string> v3 = new List<string>();

            FileInfo file = new FileInfo(@"");  //File path

            ExcelPackage Package = new ExcelPackage(file);
            //ExcelWorksheet worksheet = Package.Workbook.Worksheets.Add("Sheet 1");
            int o = Package.Workbook.Worksheets.Count;
            ExcelWorksheet worksheet = Package.Workbook.Worksheets[1];

            int v = 9;//7
            int u = 7;

            string b = worksheet.Cells["G9"].Value.ToString();

            //worksheet.Cells[i, j].ToString();
            while (true)
            {
                if (worksheet.Cells[v, u].Value == null)
                    break;
                else
                    v3.Add(worksheet.Cells[v, u].Value.ToString());
                v++;
            }

            test.excelData.Add(v3);

        }

        
    }
}

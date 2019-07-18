using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using OfficeOpenXml;
using System.Drawing;
using Microsoft.Office.Interop.Excel;

namespace GusApp
{
    public class GusValidate
    {
        public string[] Array1 { get; set; }
        public string[] Array2 { get; set; }
        string Path { get;  set; }

        public GusValidate(string[] array1_param, string[] array2_param, string pathParam)
        {
            Array1 = array1_param;
            Array2 = array2_param;
            Path = pathParam;
        }

        public void addDataToExcel()
        {
            Application excel = new Application();
            Workbook workbook = excel.Workbooks.Open(Path, ReadOnly: false, Editable: true);
            Worksheet worksheet = workbook.Worksheets.Item[1] as Worksheet;

            for (int i = 0; i < Array1.Length; i++)
            {
                Range cel1 = worksheet.Rows.Cells[i+1, 1];
                cel1.Characters[0, Array1[i].Length].Font.Color = ColorTranslator.ToOle(Color.Black);
                cel1.Value = Array1[i];

                Range cel2 = worksheet.Rows.Cells[i + 1, 2];
                cel2.Characters[0, Array2[i].Length].Font.Color = ColorTranslator.ToOle(Color.Black);
                cel2.Value = Array2[i];
            }

            excel.Application.ActiveWorkbook.Save();
            excel.Application.Quit();
            excel.Quit();
        }

        public void CompareDataFromArray()
        {
            for (int i=0;i<Array1.Length; i++)
            {
                if (Array1[i].ToLower() == Array2[i].ToLower())
                    editExcel(i);
            }
        }

        public void editExcel(int index)
        {
            Application excel = new Application();
            Workbook workbook = excel.Workbooks.Open(Path, ReadOnly: false, Editable: true);
            Worksheet worksheet = workbook.Worksheets.Item[1] as Worksheet;

            Range cel1 = worksheet.Rows.Cells[index+1, 1];
            Range cel2 = worksheet.Rows.Cells[index+1, 2];
            

            cel1.Characters[0, Array1[index].Length].Font.Color = ColorTranslator.ToOle(System.Drawing.Color.Green);
            cel2.Characters[0, Array2[index].Length].Font.Color = ColorTranslator.ToOle(System.Drawing.Color.Green);

            excel.Application.ActiveWorkbook.Save();
            excel.Application.Quit();
            excel.Quit();

        }
    }
}

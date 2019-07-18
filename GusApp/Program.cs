using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GusApp
{
    class Program
    {
        static void Main(string[] args)
        {
            string[] a1 = new string[] { "ABC", "absadasdb", "1233" };
            string[] a2 = new string[] { "abc", "addssbb", "123" };
            string path = @"C:\Users\wkret\Desktop\testowanie.xlsx";

            GusValidate gusValidate = new GusValidate(a1,a2,path);

            gusValidate.addDataToExcel();
            gusValidate.CompareDataFromArray();


        }
    }
}

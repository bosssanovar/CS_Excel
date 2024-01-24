using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfApp1
{
    internal partial class ExcelBook
    {
        private static void AddBbbSheet(XLWorkbook wb)
        {
            //ワークシートの追加
            wb.Worksheets.Add("BBB");
        }
    }
}

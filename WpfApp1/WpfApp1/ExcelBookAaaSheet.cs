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
        private void AddAaaSheet(XLWorkbook wb)
        {
            const string sheetName = "AAA";

            //ワークシートの追加
            var sheet = wb.Worksheets.Add(sheetName);


            sheet.Cell("B2").Value = aaaEntity;
            sheet.Column("B").Width = 20;
            sheet.Cell("B2").Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);

            // A4からB5までの範囲をマージ
            var range = sheet.Range("A4:B5");
            range.Merge();
            range.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);

        }
    }
}

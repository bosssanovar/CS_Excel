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
        private void AddBbbSheet(XLWorkbook wb)
        {
            const string sheetName = "BBBBB";

            //ワークシートの追加
            var sheet = wb.Worksheets.Add(sheetName);

            SetBbbContent(sheet);

            var range = sheet.Range("D2:G5");
            range.Merge();
            range.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
        }

        private void SetBbbContent(IXLWorksheet sheet)
        {
            sheet.Column("B").Width = 40;

            for (var i = 1; i <= 10; i++)
            {
                sheet.Cell($"B{i}").Value = bbbEntity;
                sheet.Cell($"B{i}").Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            }
        }
    }
}

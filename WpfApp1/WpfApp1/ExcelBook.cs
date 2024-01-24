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
        private string aaaEntity = string.Empty;
        private string bbbEntity = string.Empty;

        internal void Save(string savePath)
        {
            using var wb = new XLWorkbook();

            AddAaaSheet(wb);

            AddBbbSheet(wb);

            wb.SaveAs(savePath);
        }

        internal void SetAAAEntity(string v)
        {
            aaaEntity = v;
        }

        internal void SetBBBEntity(string v)
        {
            bbbEntity = v;
        }
    }
}

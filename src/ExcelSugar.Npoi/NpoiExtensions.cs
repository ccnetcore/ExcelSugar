using System;
using System.Collections.Generic;
using System.Text;
using NPOI.SS.UserModel;

namespace ExcelSugar.Npoi
{
    public static class NpoiExtensions
    {
        public static IRow GetOrCreateRow(this ISheet sheet, int rownum)
        {
            IRow row = sheet.GetRow(rownum);
            if (row == null)
            {
                row = sheet.CreateRow(rownum);
            }
            return row;
        }

        public static ICell GetOrCreateCell(this IRow row, int colnum)
        {
            ICell cell = row.GetCell(colnum);
            if (cell == null)
            {
                cell = row.CreateCell(colnum);
            }
            return cell;
        }
    }
}

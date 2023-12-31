﻿using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelSugar.Core.Exportable
{
    public static class ExportableExtensions
    {
        public static IOemExportable<ExcelSugarConfig> Exportable(this IExcelSugarClient excelSugarClient)
        {
            return excelSugarClient.Exportable<ExcelSugarConfig>(null);
        }

        public static IOemExportable<ExcelSugarConfig> Exportable(this IExcelSugarClient excelSugarClient,string fromPath)
        {
            return excelSugarClient.Exportable().From(fromPath);
        }
    }
}

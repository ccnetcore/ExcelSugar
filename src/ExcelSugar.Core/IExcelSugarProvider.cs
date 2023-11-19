using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelSugar.Core.Exportable;
using ExcelSugar.Core.Queryable;

namespace ExcelSugar.Core
{
    public interface IExcelSugarProvider
    {
        IOemExportable<T> CreateExportable<T>(IEnumerable<T>? expObj);
        IOemQueryable<T> CreateQueryable<T>();
    }
}

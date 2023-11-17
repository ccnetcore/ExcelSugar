using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelSugar.Core.Exportable;
using ExcelSugar.Core.Queryable;

namespace ExcelSugar.Core
{
    public interface IExcelSugarClient:IDisposable
    {
        IOemExportable<T> Exportable<T>(IEnumerable<T>? expObj);
        IOemQueryable<T> Queryable<T>();
    }
}

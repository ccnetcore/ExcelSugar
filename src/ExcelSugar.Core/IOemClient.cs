using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSugar.Core
{
    public interface IOemClient
    {
        IOemExportable<T> Exportable<T>(IEnumerable<T> expObj);
        IOemQueryable<T> Queryable<T>();
    }
}

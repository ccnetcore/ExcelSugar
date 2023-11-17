using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSugar.Core.Exportable
{
    public interface IOemExportable<T>
    {

        Task<int> ExecuteCommandAsync();
        IOemExportable<T> From(string excelPath);
    }
}

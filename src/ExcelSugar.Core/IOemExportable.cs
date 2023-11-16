using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSugar.Core
{
    public interface IOemExportable<T>
    {

        Task<int> ExecuteCommandAsync();
    }
}

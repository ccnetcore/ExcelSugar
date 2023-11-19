using System.Collections.Generic;
using ExcelSugar.Core;
using ExcelSugar.Core.Exportable;
using ExcelSugar.Core.Queryable;

namespace ExcelSugar.Npoi
{
    public class NpoiOemProvider : IExcelSugarProvider
    {
        private ExcelSugarConfig _oemConfig;
        public NpoiOemProvider(ExcelSugarConfig oemConfig)
        {
            _oemConfig = oemConfig;
        }
        public IOemQueryable<T> CreateQueryable<T>()
        {
            return new NpoiOemQueryable<T>(_oemConfig);
        }

        public IOemExportable<T> CreateExportable<T>(IEnumerable<T>? expObj)
        {
            return new NpoiOemExportable<T>(_oemConfig, expObj);
        }
    }
}

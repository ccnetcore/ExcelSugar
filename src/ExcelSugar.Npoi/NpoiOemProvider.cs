using System.Collections.Generic;
using ExcelSugar.Core;

namespace ExcelSugar.Npoi
{
    public class NpoiOemProvider : IOemProvider
    {
        private OemConfig _oemConfig;
        public NpoiOemProvider(OemConfig oemConfig)
        {
            _oemConfig = oemConfig;
        }
        public IOemQueryable<T> CreateQueryable<T>()
        {
            return new NpoiOemQueryable<T>(_oemConfig);
        }

        public IOemExportable<T> CreateExportable<T>(IEnumerable<T> expObj)
        {
            return new NpoiOemExportable<T>(_oemConfig, expObj);
        }
    }
}

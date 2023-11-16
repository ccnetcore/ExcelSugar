using System;
using System.Collections.Generic;

namespace ExcelSugar.Core
{
    public class OemClient : IOemClient
    {
        public static Func<OemConfig,IOemProvider> TestTempProviderFun;

        private OemConfig Config { get; set; }
        public OemClient(OemConfig config)
        {
            Config = config;
        }


        private IOemProvider Context => GetProvider();

        private IOemProvider GetProvider()
        {
            return TestTempProviderFun.Invoke(Config);
        }

        
        public IOemQueryable<T> Queryable<T>()
        {
            return Context.CreateQueryable<T>();
        }
        public IOemExportable<T> Exportable<T>(IEnumerable<T> expObj)
        {
            return Context.CreateExportable<T>(expObj);
        }

    }
}

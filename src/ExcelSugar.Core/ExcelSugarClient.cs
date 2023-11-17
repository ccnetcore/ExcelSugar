using System;
using System.Collections.Generic;
using System.Reflection;

namespace ExcelSugar.Core
{
    public class ExcelSugarClient : IExcelSugarClient
    {
        private IOemProvider Provider { get; set; }
        private IOemProvider Context => GetProvider();
        private IOemProvider GetProvider()
        {
            return Provider;
        }

        private OemConfig Config { get; }
        public ExcelSugarClient(OemConfig config)
        {
            Config = config;
            Init();
        }

        private void Init()
        {
            var providerFactory = new ExcelHandlerProviderFactory(Config);
            this.Provider = providerFactory.CreateOemProvider();
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

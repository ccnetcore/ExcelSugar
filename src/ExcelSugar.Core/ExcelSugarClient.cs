using System;
using System.Collections.Generic;
using System.Reflection;
using ExcelSugar.Core.Exportable;
using ExcelSugar.Core.Queryable;

namespace ExcelSugar.Core
{
    public class ExcelSugarClient : IExcelSugarClient, IDisposable
    {
        private bool disposedValue;

        private IExcelSugarProvider Provider { get; set; }
        private IExcelSugarProvider Context => GetProvider();
        private IExcelSugarProvider GetProvider()
        {
            return Provider;
        }

        private ExcelSugarConfig Config { get; }
        public ExcelSugarClient(ExcelSugarConfig config)
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
        public IOemExportable<T> Exportable<T>(IEnumerable<T>? expObj)
        {
            return Context.CreateExportable<T>(expObj);
        }

        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }
    }
}

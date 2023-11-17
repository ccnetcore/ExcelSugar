using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Threading.Tasks;
using ExcelSugar.Core.Exportable;

namespace ExcelSugar.Core
{
    public abstract class AbstractOemExportable<T> : IOemExportable<T>
    {
        protected OemConfig _config;
        protected IEnumerable<T> _expObjs;
        protected LinkedListBuilder<T> LinkedListBuilder { get; set; }
        public AbstractOemExportable(OemConfig oemConfig, IEnumerable<T>? expObj)
        {
            _config = oemConfig;
            _expObjs = expObj;
            LinkedListBuilder = new LinkedListBuilder<T>();
        }

        public abstract Task<int> ExecuteCommandAsync();

        public virtual IOemExportable<T> From(string excelPath)
        {
            LinkedListBuilder.AddFrom(excelPath);
            return this;
        }
    }
}

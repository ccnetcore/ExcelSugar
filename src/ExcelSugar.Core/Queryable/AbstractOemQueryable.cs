using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Threading.Tasks;

namespace ExcelSugar.Core.Queryable
{
    public abstract class AbstractOemQueryable<T> : IOemQueryable<T>
    {
        protected ExcelSugarConfig _config;

        protected LinkedListBuilder<T> LinkedListBuilder { get; set; }
        public AbstractOemQueryable(ExcelSugarConfig oemConfig)
        {
            _config = oemConfig;
            LinkedListBuilder = new LinkedListBuilder<T>();
        }

        public abstract Task<T> FirstAsync();

        public abstract Task<List<T>> ToListAsync();

        public virtual IOemQueryable<T> Where(Expression<Func<T, bool>> expression)
        {
            LinkedListBuilder.AddWhere(expression);
            return this;
        }

    }
}

using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Threading.Tasks;

namespace ExcelSugar.Core.Queryable
{
    public interface IOemQueryable<T>
    {
        IOemQueryable<T> Where(Expression<Func<T, bool>> expression);

        Task<List<T>> ToListAsync();

        Task<T> FirstAsync();
    }
}

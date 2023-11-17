using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Text;

namespace ExcelSugar.Core
{
    public class LinkedListBuilder<T>
    {
        /// <summary>
        /// 链表，所有数据，最终通过还会备份一个到这里，用于保存顺序
        /// </summary>
        public LinkedList<IObjectAccessor> linkedContainer = new();
        public List<ObjectAccessor<Expression<Func<T, bool>>>> WhereContainer { get; } = new();
        public List<ObjectAccessor<string>> FromContainer { get; } = new();

        public void AddWhere(Expression<Func<T, bool>> expression)
        {
            var accessor = new ObjectAccessor<Expression<Func<T, bool>>>(expression);
            WhereContainer.Add(accessor);
            linkedContainer.AddLast(accessor);
        }

        public void AddFrom(string path)
        {
            var accessor = new ObjectAccessor<string>(path);
            FromContainer.Add(accessor);
            linkedContainer.AddLast(accessor);
        }
    }
}

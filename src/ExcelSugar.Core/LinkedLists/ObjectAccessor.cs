using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelSugar.Core
{
    /// <summary>
    /// 对象访问器
    /// </summary>
    public class ObjectAccessor<T> : IObjectAccessor
    {
        private readonly T _value;
        public T Value => Get();
        public ObjectAccessor(T obj)
        {
            _value = obj;
        }

        //对象id
        public Guid Id { get; } = Guid.NewGuid();

        private T Get()
        {
            return _value;
        }
    }
}

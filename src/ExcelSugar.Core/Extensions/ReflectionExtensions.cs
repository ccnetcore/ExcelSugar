using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelSugar.Core.Extensions
{
    public static class ReflectionExtensions
    {
        public static List<T> CreateListObjct<T>()
        {
            Type listType = typeof(List<>); // 获取 List<> 的类型

            Type typeArg = typeof(T); // 指定 List<T> 中的 T 的类型
            Type constructedType = listType.MakeGenericType(typeArg); // 构造泛型类型

            List<T> result = Activator.CreateInstance(constructedType) as List<T>; // 创建 List<T> 的实例

            return result;
        }

        public static T CreateInstance<T>()
        {
            return (T)Activator.CreateInstance(typeof(T));
        }
    }
}

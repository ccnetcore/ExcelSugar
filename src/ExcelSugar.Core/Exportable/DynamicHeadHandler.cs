using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using ExcelSugar.Core.Attributes;

namespace ExcelSugar.Core.Exportable
{
    public class DynamicHeadHandler
    {

        /// <summary>
        /// 处理数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="objs"></param>
        /// <param name="dynamicHeadPropertyInfo"></param>
        /// <returns></returns>
        public List<DynamicHeadDataInfo> DataHandler<T>(DynamicHeadTypeInfo dynamicHeadTypeInfo, List<T> objs)
        {
            var results = new List<DynamicHeadDataInfo>();

            foreach (var obj in objs)
            {
                var listData = dynamicHeadTypeInfo.List.GetValue(obj);
                foreach (var modeItem in listData as IEnumerable)
                {
                    var result = new DynamicHeadDataInfo();
                    result.DataCode = dynamicHeadTypeInfo.Code.GetValue(modeItem).ToString();
                    result.DataName = dynamicHeadTypeInfo.Name.GetValue(modeItem).ToString();
                    result.DataValue = dynamicHeadTypeInfo.Value.GetValue(modeItem).ToString();
                    results.Add(result);
                }
            }
            return results;
        }

        /// <summary>
        /// 获取动态表头全部key与全部value与全部name
        /// </summary>
        /// <returns></returns>
        public DynamicHeadTypeInfo? GetDynamicHeadInfoByDynamicModel(PropertyInfo propertyInfo)
        {

            var result = new DynamicHeadTypeInfo();
            result.List = propertyInfo;
            Type? type = null;
            foreach (var interfaceType in propertyInfo.PropertyType.GetInterfaces())
            {
                if (interfaceType.IsGenericType &&
                    interfaceType.GetGenericTypeDefinition() == typeof(IEnumerable<>))
                {
                    type = interfaceType.GetGenericArguments()[0];
                    break;
                }
            }
            if (type is null)
            {
                throw new ArgumentException("参数错误，动态头需要IEnumerable");
            }

            var propertieDic = type.GetProperties().Where(x => x.GetCustomAttribute<SugarDynamicHeadAttribute>() is not null).Select(x => new KeyValuePair<SugarDynamicHeadAttribute, PropertyInfo>(x.GetCustomAttribute<SugarDynamicHeadAttribute>(), x)).ToList();

            foreach (var kv in propertieDic)
            {
                if (kv.Key.IsValue == true)
                {
                    result.Value = kv.Value;
                }
                else if (kv.Key.IsCode == true)
                {
                    result.Code = kv.Value;
                }
                else if (kv.Key.IsName == true)
                {
                    result.Name = kv.Value;
                }
            }
            result.Verify();
            return result;
        }

        /// <summary>
        /// 获取动态表头类型
        /// </summary
        /// <returns></returns>
        public PropertyInfo? GetDynamicHeadPropertyInfoByModel(Type type)
        {
            var resultProperties = type.GetProperties().Where(x => x.GetCustomAttribute<SugarDynamicHeadAttribute>() is not null).FirstOrDefault();
            if (resultProperties is null)
            {
                return null;
            }
            else
            {
                return resultProperties;
            }
        }

        /// <summary>
        /// 根据模型类获取动态表头属性类型
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public DynamicHeadTypeInfo? GetDynamicHeadInfoByModel(Type type)
        {
            var dynamicHeadModelType = GetDynamicHeadPropertyInfoByModel(type);
            if (dynamicHeadModelType is null)
            {
                return null;
            }
            return GetDynamicHeadInfoByDynamicModel(dynamicHeadModelType);
        }
    }
}

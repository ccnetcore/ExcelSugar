using System;
using System.Linq;
using System.Reflection;
using ExcelSugar.Core.Attributes;
using ExcelSugar.Core.Exportable;

namespace ExcelSugar.Core.Extensions
{
    public static class AttributeExtensions
    {
        public static PropertyInfo[] GetValidProperties(this Type type)
        {
            return type.GetProperties().Where(x => x.GetCustomAttribute<SugarHeadAttribute>() is not null).Where(x => x.GetGetMethod().IsPublic).ToArray();
        }

        public static string GetSheetNameFromType(this Type type)
        {
            var sheetName = type.GetCustomAttribute<SugarSheetAttribute>()?.SheetName;
            return sheetName ?? type.Name;

        }



    }
}

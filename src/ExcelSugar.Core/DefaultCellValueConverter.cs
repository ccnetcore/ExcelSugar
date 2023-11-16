using System;
using System.Collections.Generic;
using System.Text;
using System.Text.Json;

namespace ExcelSugar.Core
{

    /// <summary>
    /// 单元格值转换器
    /// </summary>
    public class DefaultCellValueConverter: ICellValueConverter
    {

        /// <summary>
        /// 单元格的值转换未对象属性值
        /// </summary>
        /// <param name="cellValue"></param>
        /// <param name="propertyType"></param>
        /// <returns></returns>
        public object? CellToProperty(string cellValue, Type propertyType)
        {
            object value = cellValue.TrimStart("\"".ToCharArray()).TrimEnd("\"".ToCharArray());

            return Convert.ChangeType(cellValue, propertyType);
        }

        /// <summary>
        /// 属性值转单元格的值
        /// </summary>
        /// <param name="propertyValue"></param>
        /// <returns></returns>
        public string PropertyToCell(object? propertyValue)
        {
            return JsonSerializer.Serialize(propertyValue).TrimStart("\"".ToCharArray()).TrimEnd("\"".ToCharArray());
        }
    }
}

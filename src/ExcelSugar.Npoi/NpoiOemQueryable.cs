using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Threading.Tasks;
using ExcelSugar.Core;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelSugar.Npoi
{
    public class NpoiOemQueryable<T> : IOemQueryable<T>
    {
        private OemConfig _config;
        public NpoiOemQueryable(OemConfig oemConfig)
        {
            _config = oemConfig;
        }

        public Task<T> FirstAsync()
        {
            throw new NotImplementedException();
        }

        public Task<List<T>> ToListAsync()
        {
            return Task.FromResult(Import<T>(_config.Path));
        }

        public IOemQueryable<T> Where(Expression<Func<T, bool>> expression)
        {
            throw new NotImplementedException();
        }





        /// <summary>
        /// 导入excel
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="filePath"></param>
        /// <returns></returns>
        private List<T> Import<T>(string filePath)
        {
            Type listType = typeof(List<>); // 获取 List<> 的类型

            Type typeArg = typeof(T); // 指定 List<T> 中的 T 的类型
            Type constructedType = listType.MakeGenericType(typeArg); // 构造泛型类型

            List<T> result = Activator.CreateInstance(constructedType) as List<T>; // 创建 List<T> 的实例


            Dictionary<int, PropertyInfo> propHas = new Dictionary<int, PropertyInfo>();
            var properties = typeof(T).GetProperties().Where(x => x.GetCustomAttribute<DisplayNameAttribute>() is not null).Where(x => x.GetGetMethod().IsPublic).ToList();


            // 创建文件流
            using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                // 创建工作簿
                IWorkbook workbook = new XSSFWorkbook(fileStream);

                var sheetName = typeof(T).GetCustomAttribute<SugarSheetAttribute>()?.SheetName;
                var sheetReplaceName = sheetName ?? typeof(T).Name;
                // 选择第一个工作表
                ISheet sheet = workbook.GetSheet(sheetReplaceName);

                //获取表头
                IRow headerRow = sheet.GetRow(0);


                for (int col = 0; col < headerRow.LastCellNum; col++)
                {
                    // 获取单元格的值
                    ICell cell = headerRow.GetCell(col);
                    var property = properties.Where(x => x.GetCustomAttribute<DisplayNameAttribute>().DisplayName == cell.StringCellValue).FirstOrDefault();
                    if (property is not null && !propHas.Values.Contains(property))
                    {
                        propHas[col] = property;
                    }
                }

                // 遍历行
                for (int row = 1; row <= sheet.LastRowNum; row++)
                {
                    //一行一个对象
                    IRow currentRow = sheet.GetRow(row);
                    var currentResult = (T)Activator.CreateInstance(typeArg);
                    // 遍历列
                    for (int col = 0; col < currentRow.LastCellNum; col++)
                    {
                        // 获取单元格的值
                        ICell cell = currentRow.GetCell(col);
                        string? cellValue = cell.ToString();
                        object value = cellValue.TrimStart("\"".ToCharArray()).TrimEnd("\"".ToCharArray());

                        value = Convert.ChangeType(cellValue, propHas[col].PropertyType);


                        propHas[col].SetValue(currentResult, value);

                    }

                    result.Add(currentResult);
                }
            }

            return result;
        }

    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Threading.Tasks;
using ExcelSugar.Core;
using ExcelSugar.Core.Extensions;
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
            return Task.FromResult(Import(_config.Path));
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
        private List<T> Import(string filePath)
        {
            List<T> result = ReflectionExtensions.CreateListObjct<T>();

            Dictionary<int, PropertyInfo> propHas = new Dictionary<int, PropertyInfo>();
            var properties = typeof(T).GetValidProperties();

            // 创建文件流
            using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                // 创建工作簿
                IWorkbook workbook = new XSSFWorkbook(fileStream);

                var sheetName = typeof(T).GetSheetNameFromType();
                // 选择第一个工作表
                ISheet sheet = workbook.GetSheet(sheetName);

                //获取表头
                IRow headerRow = sheet.GetRow(0);


                for (int col = 0; col < headerRow.LastCellNum; col++)
                {
                    // 获取单元格的值
                    ICell cell = headerRow.GetCell(col);
                    var property = properties.Where(x => x.GetCustomAttribute<SugarHeadAttribute>().DisplayName == cell.StringCellValue).FirstOrDefault();
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
                    var currentResult = ReflectionExtensions.CreateInstance<T>();
                    // 遍历列
                    for (int col = 0; col < currentRow.LastCellNum; col++)
                    {
                        // 获取单元格的值
                        ICell cell = currentRow.GetCell(col);
                        string? cellValue = cell.ToString();

                        //属性转换值
                        object? value = _config.CeellValueConverter.CellToProperty(cellValue, propHas[col].PropertyType);

                        propHas[col].SetValue(currentResult, value);

                    }

                    result.Add(currentResult);
                }
            }

            return result;
        }

    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Threading.Tasks;
using ExcelSugar.Core;
using ExcelSugar.Core.Extensions;
using ExcelSugar.Core.Queryable;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelSugar.Npoi
{
    public class NpoiOemQueryable<T> : AbstractOemQueryable<T>
    {

        public NpoiOemQueryable(OemConfig oemConfig) : base(oemConfig)
        {
        }

        public override Task<T> FirstAsync()
        {
            throw new NotImplementedException();
        }

        public override Task<List<T>> ToListAsync()
        {
            return Task.FromResult(Import(_config.Path));
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

                    var whereResult = true;
                    //先去判断表达式，该行是否成立
                    foreach (var whereExp in base.LinkedListBuilder.WhereContainer)
                    {
                        var expression = (MemberExpression)((BinaryExpression)whereExp.Value.Body).Left;
                        var propertieName = expression.Member.Name;
                        var oropertyInfo = typeof(T).GetProperty(propertieName);

                        //值是excel中获取到的,获取改行的关系
                        var currentColKv = propHas.Where(x => x.Value.Name == propertieName).FirstOrDefault();
                        // 获取单元格的值
                        ICell cell = currentRow.GetCell(currentColKv.Key);
                        string? cellValue = cell.ToString();
                        //属性转换值
                        object? value = _config.CellValueConverter.CellToProperty(cellValue, currentColKv.Value.PropertyType);

                        //给创建的对象赋值
                        oropertyInfo.SetValue(currentResult, value);
                        var whereFunc = whereExp.Value.Compile();
                        if (whereFunc.Invoke(currentResult) == false)
                        {
                            whereResult = false;
                            break;
                        };
                    }

                    //不满住where条件，跳过该行
                    if (whereResult == false)
                    {
                        continue;
                    }

                    // 遍历列
                    for (int col = 0; col < currentRow.LastCellNum; col++)
                    {
                        //如果列不在已定义中的，直接跳过即可
                        if (!propHas.ContainsKey(col))
                        {
                            continue;
                        }



                        // 获取单元格的值
                        ICell cell = currentRow.GetCell(col);
                        string? cellValue = cell.ToString();

                        //属性转换值
                        object? value = _config.CellValueConverter.CellToProperty(cellValue, propHas[col].PropertyType);

                        propHas[col].SetValue(currentResult, value);

                    }

                    result.Add(currentResult);
                }
            }

            return result;
        }

    }
}

using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using ExcelSugar.Core;
using ExcelSugar.Core.Extensions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelSugar.Npoi
{
    public class NpoiOemExportable<T> : IOemExportable<T>
    {
        private OemConfig _config;
        private IEnumerable<T> _expObjs;

        public NpoiOemExportable(OemConfig oemConfig, IEnumerable<T> expObj)
        {
            _config = oemConfig;
            _expObjs = expObj;
        }
        private IOemExportable<T> From(string excelPath)
        {
            return this;
        }

        public Task<int> ExecuteCommandAsync()
        {
            Export(_expObjs.ToList(), _config.Path);
            return Task.FromResult(1);
        }

        /// <summary>
        /// 导出excel
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="entityList"></param>
        /// <param name="filePath"></param>
        private void Export(List<T> entityList, string filePath)
        {
            var properties = typeof(T).GetValidProperties();
            // 创建工作簿
            IWorkbook workbook = new XSSFWorkbook();

            var sheetName = typeof(T).GetSheetNameFromType();
            // 创建工作表
            ISheet sheet = workbook.CreateSheet(sheetName);


            // 写入表头
            IRow headerRow = sheet.CreateRow(0);
            for (int j = 0; j < properties.Count(); j++)
            {
                headerRow.CreateCell(j).SetCellValue(properties[j].GetCustomAttribute<SugarHeadAttribute>()!.DisplayName);
            }

            // 写入数据
            for (int i = 0; i < entityList.Count(); i++)
            {
                var currentEntity = entityList[i];
                IRow dataRow = sheet.CreateRow(i + 1);
                for (int j = 0; j < properties.Count(); j++)
                {
                    var currentPropertiy = properties[j];
                    var cellStr = _config.CellValueConverter.PropertyToCell(currentPropertiy.GetValue(currentEntity));
                    //只处理简单类型
                    dataRow.CreateCell(j).SetCellValue(cellStr);
                }
            }

            // 保存文件
            using (FileStream fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
            }

            workbook.Dispose();
        }


    }
}

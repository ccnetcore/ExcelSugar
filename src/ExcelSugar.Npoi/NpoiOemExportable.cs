using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.Json;
using System.Threading.Tasks;
using ExcelSugar.Core;
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
        private void Export<T>(List<T> entityList, string filePath)
        {
            var properties = typeof(T).GetProperties().Where(x => x.GetCustomAttribute<DisplayNameAttribute>() is not null).Where(x => x.GetGetMethod().IsPublic).ToList();
            // 创建工作簿
            IWorkbook workbook = new XSSFWorkbook();
            // 创建工作表
            ISheet sheet = workbook.CreateSheet("sheet");


            // 写入表头
            IRow headerRow = sheet.CreateRow(0);
            for (int j = 0; j < properties.Count(); j++)
            {
                headerRow.CreateCell(j).SetCellValue(properties[j].GetCustomAttribute<DisplayNameAttribute>()!.DisplayName);
            }

            // 写入数据
            for (int i = 0; i < entityList.Count(); i++)
            {
                var currentEntity = entityList[i];
                IRow dataRow = sheet.CreateRow(i + 1);
                for (int j = 0; j < properties.Count(); j++)
                {
                    var currentPropertiy = properties[j];

                    //只处理简单类型
                    dataRow.CreateCell(j).SetCellValue(JsonSerializer.Serialize(currentPropertiy.GetValue(currentEntity)).TrimStart("\"".ToCharArray()).TrimEnd("\"".ToCharArray()));
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

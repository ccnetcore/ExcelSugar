using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelSugar.Core;

namespace ExcelSugar.Npoi.Test
{
    public class NpoiTestBase
    {

        protected IExcelSugarClient CreateClient()
        {
            //创建客户端
            IExcelSugarClient excelSugarClient = new ExcelSugarClient(new ExcelSugarConfig { Path = "../../../TempExcel/Test.xlsx", HandlerType = ExcelHandlerType.Npoi });
            //支持动态列模型
            return excelSugarClient;
        }

        protected List<TestModel> CreateNullTestModel()
        {
            return new List<TestModel>();
        }

        protected List<TestModel> CreateTestModel()
        {
            var entities = new List<TestModel> {
                new TestModel { Description = "男的", Name = "张三", DynamicModels=new List<DynamicModel>{
                    new DynamicModel { DataCode="height",DataName="身高",DataValue=188},
                    new DynamicModel { DataCode="age",DataName="年龄",DataValue=35},
                } },
                new TestModel { Description = "女的", Name = "李四" ,DynamicModels=new List<DynamicModel>{
                    new DynamicModel { DataCode="height",DataName="身高",DataValue=168},
                    new DynamicModel { DataCode="age",DataName="年龄",DataValue=18},
                } }
            };
            return entities;
        }
    }
}

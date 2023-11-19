using ExcelSugar.Core;
using ExcelSugar.Core.Attributes;
using Xunit;
using Xunit.Abstractions;

namespace ExcelSugar.Npoi.Test
{
    public class Npoi_Test
    {
        private ITestOutputHelper _output;
        public Npoi_Test(ITestOutputHelper output)
        {
            _output = output;

        }
       

        [Fact]
        public async Task Test1()
        {
            //创建客户端
            IExcelSugarClient excelSugarClient = new ExcelSugarClient(new OemConfig { Path = "../../../Test3.xlsx", HandlerType = ExcelHandlerType.Npoi });
            //创建实体
            var entities = new List<TestModel> {
                new TestModel { Description = "111", Name = "222", DynamicModels=new List<DynamicModel>{ 
                    new DynamicModel { DataCode="a1",DataName="测试头",DataValue=123},
                new DynamicModel { DataCode="a2",DataName="测试头2",DataValue=456},

                } },
                new TestModel { Description = "333", Name = "444" ,DynamicModels=new List<DynamicModel>{
                    
                    new DynamicModel { DataCode="a2",DataName="测试头2",DataValue=2223},
                 new DynamicModel { DataCode="a1",DataName="测试头",DataValue=33331},

                } }
            };

            //导出：传入实体对象,返回excel文件
            await excelSugarClient.Exportable(entities).ExecuteCommandAsync();

            ////导出：来自模板，传入实体对象,给一个from模板路径，根据模板的样式返回excel文件
            //await excelSugarClient.Exportable(entities).From("../../../Test.xlsx").ExecuteCommandAsync();

            ////导出：返回一个空模板
            //await excelSugarClient.Exportable("../../../Test.xlsx").ExecuteCommandAsync();

            ////查询：从excel中查询，返回实体对象
            //var data = await excelSugarClient.Queryable<TestModel>().ToListAsync();

            ////查询：从excel中查询，包含条件，返回实体对象
            //var data2 = await excelSugarClient.Queryable<TestModel>().Where(x => x.Name == "张三").ToListAsync();

            ////释放对象
            //excelSugarClient.Dispose();
            //Assert.True(data.Any());
        }
    }


    [SugarSheet("测试")]
    class TestModel
    {
        [SugarHead("姓名")]
        public string Name { get; set; }
        [SugarHead("描述")]
        public string Description { get; set; }

        /// <summary>
        /// 包含动态头类型
        /// </summary>
        [SugarDynamicHead]
        public List<DynamicModel> DynamicModels { get; set; } = new List<DynamicModel>();
        //[SugarHead("对象", IsJson = true)]
        //public Dictionary<int, string> KKK { get; set; } = new Dictionary<int, string>();
    }
    /// <summary>
    /// 动态头类型中，必须要包含code、value、name
    /// </summary>
    class DynamicModel
    {
        [SugarDynamicHead(IsCode = true)]
        public string DataCode { get; set; }

        [SugarDynamicHead(IsValue = true)]
        public int DataValue { get; set; }

        [SugarDynamicHead(IsName = true)]
        public string DataName { get; set; }
    }
}
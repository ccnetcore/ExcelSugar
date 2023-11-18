using System.ComponentModel;
using ExcelSugar.Core;
using Xunit;
using Xunit.Abstractions;
using ExcelSugar.Core.Exportable;

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
            _output.WriteLine("这是一条测试消息。");

            //创建客户端
            IExcelSugarClient excelSugarClient = new ExcelSugarClient(new OemConfig { Path = "../../../Test2.xlsx", HandlerType = ExcelHandlerType.Npoi });

            //创建实体
            var entities = new List<TestModel> {
                new TestModel { Description = "男的", Name = "张三" },
                new TestModel { Description = "女的", Name = "李四" }
            };
            //查询：从excel中查询，包含条件，返回实体对象
            var data3 = await excelSugarClient.Queryable<TestModel>().Where(x => x.Name == "张三").ToListAsync();
            { 
            }


            //导出：传入实体对象,返回excel文件
            await excelSugarClient.Exportable(entities).ExecuteCommandAsync();

            //导出：来自模板，传入实体对象,给一个from模板路径，根据模板的样式返回excel文件
            await excelSugarClient.Exportable(entities).From("../../../Test.xlsx").ExecuteCommandAsync();

            //导出：返回一个空模板
            await excelSugarClient.Exportable("../../../Test.xlsx").ExecuteCommandAsync();

            //查询：从excel中查询，返回实体对象
            var data = await excelSugarClient.Queryable<TestModel>().ToListAsync();

            //查询：从excel中查询，包含条件，返回实体对象
            var data2 = await excelSugarClient.Queryable<TestModel>().Where(x=>x.Name=="张三").ToListAsync();

            //释放对象
            excelSugarClient.Dispose();
            Assert.True(data.Any());
        }
    }


    [SugarSheet("测试")]
    class TestModel
    {
        [SugarHead("姓名")]
        public string Name { get; set; }
        [SugarHead("描述")]
        public string Description { get; set; }

        //[SugarHead("对象", IsJson = true)]
        //public Dictionary<int, string> KKK { get; set; } = new Dictionary<int, string>();
    }
}
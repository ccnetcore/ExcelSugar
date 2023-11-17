using System.ComponentModel;
using ExcelSugar.Core;
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
            _output.WriteLine("����һ��������Ϣ��");

            IExcelSugarClient oemClient = new ExcelSugarClient(new OemConfig { Path = "../../../Test.xlsx", HandlerType= ExcelHandlerType.Npoi});

            var sss = new List<TestModel> {
                new TestModel { Description = "123", Name = "123" },
                new TestModel { Description = "3124", Name = "312412" }
            };

            await oemClient.Exportable(sss).ExecuteCommandAsync();
            var data = await oemClient.Queryable<TestModel>().ToListAsync();


            Assert.True(data.Any());
        }
    }


    [SugarSheet("����")]
    class TestModel
    {
        [SugarHead("����")]
        public string Name { get; set; }
        [SugarHead("����")]
        public string Description { get; set; }

        //[SugarHead("����", IsJson = true)]
        //public Dictionary<int, string> KKK { get; set; } = new Dictionary<int, string>();
    }
}
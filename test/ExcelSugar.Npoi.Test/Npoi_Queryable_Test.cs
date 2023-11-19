using ExcelSugar.Core;
using ExcelSugar.Core.Attributes;
using Xunit;
using Xunit.Abstractions;

namespace ExcelSugar.Npoi.Test
{
    public class Npoi_Exportable_Test : NpoiTestBase
    {
        private ITestOutputHelper _output;
        public Npoi_Exportable_Test(ITestOutputHelper output)
        {
            _output = output;

        }

        /// <summary>
        /// 根据excel查询全部
        /// </summary>
        /// <returns></returns>
        [Fact]
        public async Task Queryable_Test()
        {
            var client = CreateClient();
            var data = await client.Queryable<TestModel>().ToListAsync();
            client.Dispose();
            Assert.True(data.Count == 2);
        }

        /// <summary>
        /// 根据excel查询并where表达式筛选，非全读，效率较高
        /// </summary>
        /// <returns></returns>
        [Fact]
        public async Task Queryable_Where_Test()
        {
            var client = CreateClient();
            var data = await client.Queryable<TestModel>().Where(x => x.Name == "张三").ToListAsync();
            client.Dispose();
            Assert.True(data.Count == 1);
        }
    }
}
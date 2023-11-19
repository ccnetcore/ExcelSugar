using ExcelSugar.Core;
using ExcelSugar.Core.Attributes;
using Xunit;
using Xunit.Abstractions;

namespace ExcelSugar.Npoi.Test
{
    public class Npoi_Queryable_Test : NpoiTestBase
    {
        private ITestOutputHelper _output;
        public Npoi_Queryable_Test(ITestOutputHelper output)
        {
            _output = output;

        }

        /// <summary>
        /// 直接导出创建的模型，支持动态表头
        /// </summary>
        /// <returns></returns>
        [Fact]
        public async Task Exportable_Test()
        {
            var client = CreateClient();
            var testModel = CreateTestModel();
            await client.Exportable(testModel).ExecuteCommandAsync();
            client.Dispose();
            Assert.True(true);
        }


        /// <summary>
        /// 空测试
        /// </summary>
        /// <returns></returns>
        [Fact]
        public async Task Exportable_Null_Test()
        {
            var client = CreateClient();
            var testModel = CreateNullTestModel();
            await client.Exportable(testModel).ExecuteCommandAsync();
            client.Dispose();
            Assert.True(true);
        }
   
        /// <summary>
        /// 根据提供的模板进行导出创建的模型
        /// </summary>
        /// <returns></returns>
        [Fact]
        public async Task Exportable_From_Test()
        {
            var client = CreateClient();
            var testModel = CreateTestModel();
            await client.Exportable(testModel).From("../../../TempExcel/Template.xlsx").ExecuteCommandAsync();
            //Or await client.Exportable("../../../Test.xlsx").ExecuteCommandAsync();
            client.Dispose();
            Assert.True(true);
        }

        /// <summary>
        /// 空测试
        /// </summary>
        /// <returns></returns>
        [Fact]
        public async Task Exportable_From_Null_Test()
        {
            var client = CreateClient();
            var testModel = CreateNullTestModel();
            await client.Exportable(testModel).From("../../../TempExcel/Template.xlsx").ExecuteCommandAsync();
            //Or await client.Exportable("../../../Test.xlsx").ExecuteCommandAsync();
            client.Dispose();
            Assert.True(true);
        }


        /// <summary>
        /// 直接导出模板
        /// </summary>
        /// <returns></returns>
        [Fact]
        public async Task Exportable_Template_Test()
        {
            var client = CreateClient();
            await client.Exportable("../../../Test.xlsx").ExecuteCommandAsync();
            client.Dispose();
            Assert.True(true);
        }
    }



}
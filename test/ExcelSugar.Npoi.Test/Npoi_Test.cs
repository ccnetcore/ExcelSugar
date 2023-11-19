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

            //鍒涘缓瀹㈡埛绔?
            IExcelSugarClient excelSugarClient = new ExcelSugarClient(new OemConfig { Path = "../../../Test3.xlsx", HandlerType = ExcelHandlerType.Npoi });
            //鍒涘缓瀹炰綋

            var entities = new List<TestModel> {
                new TestModel { Description = "你好呀", Name = "222", DynamicModels=new List<DynamicModel>{
                    new DynamicModel { DataCode="a1",DataName="娴嬭瘯澶",DataValue=123},
                new DynamicModel { DataCode="a2",DataName="娴嬭瘯澶?",DataValue=456},

                } },
                new TestModel { Description = "333", Name = "444" ,DynamicModels=new List<DynamicModel>{

                    new DynamicModel { DataCode="a2",DataName="娴嬭瘯澶?",DataValue=2223},
                 new DynamicModel { DataCode="a1",DataName="娴嬭瘯澶?",DataValue=33331},

                } }
            };
            ////查询：从excel中查询，包含条件，返回实体对象
            //var data3 = await excelSugarClient.Queryable<TestModel>().Where(x => x.Name == "张三").ToListAsync();
            //{
            //}


            //瀵煎嚭锛氫紶鍏ュ疄浣撳璞?杩斿洖excel鏂囦欢
            await excelSugarClient.Exportable(entities).ExecuteCommandAsync();

            ////瀵煎嚭锛氭潵鑷ā鏉匡紝浼犲叆瀹炰綋瀵硅薄,缁欎竴涓猣rom妯℃澘璺緞锛屾牴鎹ā鏉跨殑鏍峰紡杩斿洖excel鏂囦欢
            //await excelSugarClient.Exportable(entities).From("../../../Test.xlsx").ExecuteCommandAsync();

            ////瀵煎嚭锛氳繑鍥炰竴涓┖妯℃澘
            //await excelSugarClient.Exportable("../../../Test.xlsx").ExecuteCommandAsync();

            ////鏌ヨ锛氫粠excel涓煡璇紝杩斿洖瀹炰綋瀵硅薄
            //var data = await excelSugarClient.Queryable<TestModel>().ToListAsync();

            ////鏌ヨ锛氫粠excel涓煡璇紝鍖呭惈鏉′欢锛岃繑鍥炲疄浣撳璞?
            //var data2 = await excelSugarClient.Queryable<TestModel>().Where(x => x.Name == "寮犱笁").ToListAsync();

            ////閲婃斁瀵硅薄
            //excelSugarClient.Dispose();
            //Assert.True(data.Any());
        }
    }


    [SugarSheet("娴嬭瘯")]
    class TestModel
    {
        [SugarHead("濮撳悕")]
        public string Name { get; set; }
        [SugarHead("鎻忚堪")]
        public string Description { get; set; }

        /// <summary>
        /// 鍖呭惈鍔ㄦ€佸ご绫诲瀷
        /// </summary>
        [SugarDynamicHead]
        public List<DynamicModel> DynamicModels { get; set; } = new List<DynamicModel>();
        //[SugarHead("瀵硅薄", IsJson = true)]
        //public Dictionary<int, string> KKK { get; set; } = new Dictionary<int, string>();
    }
    /// <summary>
    /// 鍔ㄦ€佸ご绫诲瀷涓紝蹇呴』瑕佸寘鍚玞ode銆乿alue銆乶ame
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
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
            //�����ͻ���
            IExcelSugarClient excelSugarClient = new ExcelSugarClient(new OemConfig { Path = "../../../Test3.xlsx", HandlerType = ExcelHandlerType.Npoi });
            //����ʵ��
            var entities = new List<TestModel> {
                new TestModel { Description = "111", Name = "222", DynamicModels=new List<DynamicModel>{ 
                    new DynamicModel { DataCode="a1",DataName="����ͷ",DataValue=123},
                new DynamicModel { DataCode="a2",DataName="����ͷ2",DataValue=456},

                } },
                new TestModel { Description = "333", Name = "444" ,DynamicModels=new List<DynamicModel>{
                    
                    new DynamicModel { DataCode="a2",DataName="����ͷ2",DataValue=2223},
                 new DynamicModel { DataCode="a1",DataName="����ͷ",DataValue=33331},

                } }
            };
            //��ѯ����excel�в�ѯ����������������ʵ�����
            var data3 = await excelSugarClient.Queryable<TestModel>().Where(x => x.Name == "����").ToListAsync();
            { 
            }


            //����������ʵ�����,����excel�ļ�
            await excelSugarClient.Exportable(entities).ExecuteCommandAsync();

            ////����������ģ�壬����ʵ�����,��һ��fromģ��·��������ģ�����ʽ����excel�ļ�
            //await excelSugarClient.Exportable(entities).From("../../../Test.xlsx").ExecuteCommandAsync();

            ////����������һ����ģ��
            //await excelSugarClient.Exportable("../../../Test.xlsx").ExecuteCommandAsync();

            ////��ѯ����excel�в�ѯ������ʵ�����
            //var data = await excelSugarClient.Queryable<TestModel>().ToListAsync();

            ////��ѯ����excel�в�ѯ����������������ʵ�����
            //var data2 = await excelSugarClient.Queryable<TestModel>().Where(x => x.Name == "����").ToListAsync();

            ////�ͷŶ���
            //excelSugarClient.Dispose();
            //Assert.True(data.Any());
        }
    }


    [SugarSheet("����")]
    class TestModel
    {
        [SugarHead("����")]
        public string Name { get; set; }
        [SugarHead("����")]
        public string Description { get; set; }

        /// <summary>
        /// ������̬ͷ����
        /// </summary>
        [SugarDynamicHead]
        public List<DynamicModel> DynamicModels { get; set; } = new List<DynamicModel>();
        //[SugarHead("����", IsJson = true)]
        //public Dictionary<int, string> KKK { get; set; } = new Dictionary<int, string>();
    }
    /// <summary>
    /// ��̬ͷ�����У�����Ҫ����code��value��name
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
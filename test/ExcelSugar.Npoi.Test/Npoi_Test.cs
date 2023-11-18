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
            _output.WriteLine("����һ��������Ϣ��");

            //�����ͻ���
            IExcelSugarClient excelSugarClient = new ExcelSugarClient(new OemConfig { Path = "../../../Test2.xlsx", HandlerType = ExcelHandlerType.Npoi });

            //����ʵ��
            var entities = new List<TestModel> {
                new TestModel { Description = "�е�", Name = "����" },
                new TestModel { Description = "Ů��", Name = "����" }
            };
            //��ѯ����excel�в�ѯ����������������ʵ�����
            var data3 = await excelSugarClient.Queryable<TestModel>().Where(x => x.Name == "����").ToListAsync();
            { 
            }


            //����������ʵ�����,����excel�ļ�
            await excelSugarClient.Exportable(entities).ExecuteCommandAsync();

            //����������ģ�壬����ʵ�����,��һ��fromģ��·��������ģ�����ʽ����excel�ļ�
            await excelSugarClient.Exportable(entities).From("../../../Test.xlsx").ExecuteCommandAsync();

            //����������һ����ģ��
            await excelSugarClient.Exportable("../../../Test.xlsx").ExecuteCommandAsync();

            //��ѯ����excel�в�ѯ������ʵ�����
            var data = await excelSugarClient.Queryable<TestModel>().ToListAsync();

            //��ѯ����excel�в�ѯ����������������ʵ�����
            var data2 = await excelSugarClient.Queryable<TestModel>().Where(x=>x.Name=="����").ToListAsync();

            //�ͷŶ���
            excelSugarClient.Dispose();
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
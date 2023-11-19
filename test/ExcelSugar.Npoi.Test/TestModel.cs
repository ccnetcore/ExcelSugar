using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelSugar.Core.Attributes;
using ExcelSugar.Core;

namespace ExcelSugar.Npoi.Test
{

    [SugarSheet("测试表")]
    public class TestModel
    {
        [SugarHead("姓名")]
        public string Name { get; set; }
        [SugarHead("描述")]
        public string Description { get; set; }

        /// <summary>
        /// 导出可支持动态列模型
        /// </summary>
        [SugarDynamicHead]
        public List<DynamicModel> DynamicModels { get; set; } = new List<DynamicModel>();

        //正在支持中。。。。。。理论上支持
        //[SugarHead("对象列", IsJson = true)]
        //public Dictionary<int, string> KKK { get; set; } = new Dictionary<int, string>();
    }
    /// <summary>
    /// 动态表头，需具备IsCode、IsValue、IsName 3个字段
    /// </summary>
    public class DynamicModel
    {
        [SugarDynamicHead(IsCode = true)]
        public string DataCode { get; set; }

        [SugarDynamicHead(IsValue = true)]
        public int DataValue { get; set; }

        [SugarDynamicHead(IsName = true)]
        public string DataName { get; set; }
    }

}

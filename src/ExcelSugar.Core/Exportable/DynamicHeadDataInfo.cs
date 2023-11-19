using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelSugar.Core.Exportable
{

    /// <summary>
    /// 行
    /// </summary>
    public class DynamicHeadDataRowInfo : List<DynamicHeadDataInfo>
    {


    }

    public class DynamicHeadDataInfo
    {
        public string? DataCode { get; set; }
        public string? DataName { get; set; }
        public string? DataValue { get; set; }
    }
}

using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace ExcelSugar.Core.Exportable
{
    public class DataCodeComparer : IEqualityComparer<DynamicHeadDataInfo>
    {
        bool IEqualityComparer<DynamicHeadDataInfo>.Equals(DynamicHeadDataInfo x, DynamicHeadDataInfo y)
        {
            return x.DataCode == y.DataCode;
        }

        int IEqualityComparer<DynamicHeadDataInfo>.GetHashCode(DynamicHeadDataInfo obj)
        {
            return obj.DataCode.GetHashCode();
        }
    }
}

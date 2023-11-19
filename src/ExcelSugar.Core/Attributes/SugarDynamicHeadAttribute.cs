using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelSugar.Core.Attributes
{
    public class SugarDynamicHeadAttribute : Attribute
    {
        public bool IsCode { get; set; }
        public bool IsValue { get; set; }

        public bool IsName { get; set; }
        public SugarDynamicHeadAttribute()
        {

        }
    }
}

using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelSugar.Core
{
    public class SugarSheetAttribute:Attribute
    {
        public string? SheetName { get; set; }

        public SugarSheetAttribute()
        { 
        
        }

        public SugarSheetAttribute(string sheetName)
        {
            SheetName=sheetName;
        }
    }
}

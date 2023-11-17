using System.ComponentModel;

namespace ExcelSugar.Core
{
    public class SugarHeadAttribute : DisplayNameAttribute
    {
        public bool IsJson = false;
        public SugarHeadAttribute(string displayName) : base(displayName) { }
    }
}

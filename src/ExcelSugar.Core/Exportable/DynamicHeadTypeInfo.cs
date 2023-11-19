using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace ExcelSugar.Core.Exportable
{
    public class DynamicHeadTypeInfo
    {


        public PropertyInfo? List { get; set; }
        public PropertyInfo? Code { get; set; }
        public PropertyInfo? Name { get; set; }
        public PropertyInfo? Value { get; set; }

        public void Verify()
        {
            if (List is null)
            {
                throw new ArgumentNullException(nameof(List));
            }
            if (Code is null)
            { 
            throw new ArgumentNullException(nameof(Code));
            }
            if (Name is null)
            {
                throw new ArgumentNullException(nameof(Name));
            }
            if (Value is null)
            {
                throw new ArgumentNullException(nameof(Value));
            }

        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSugar.Core
{
    public class OemConfig
    {
        public string Path { get; set; }

        public ICellValueConverter CeellValueConverter { get; set; } = new DefaultCellValueConverter();
    }
}

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

        public ICellValueConverter CellValueConverter { get; set; } = new DefaultCellValueConverter();

        public ExcelHandlerType HandlerType { get; set; }
    }
}

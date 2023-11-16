using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelSugar.Core
{
    public interface ICellValueConverter
    {
        object? CellToProperty(string cellValue, Type propertyType);
        string PropertyToCell(object? propertyValue);
    }
}

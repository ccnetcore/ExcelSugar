namespace ExcelSugar.Core
{
    public class ExcelSugarConfig
    {
        public string Path { get; set; }

        public ICellValueConverter CellValueConverter { get; set; } = new DefaultCellValueConverter();

        public ExcelHandlerType HandlerType { get; set; } = ExcelHandlerType.Npoi;
    }
}

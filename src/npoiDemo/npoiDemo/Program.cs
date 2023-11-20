// See https://aka.ms/new-console-template for more information
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

Console.WriteLine("Hello, World!");

try
{
    using (FileStream fileStream = new FileStream("../../../2023-11-20 16-42-33模组配置表 (2).xlsx", FileMode.Open, FileAccess.Read))
    {
        // 创建工作簿
        IWorkbook workbook = new XSSFWorkbook(fileStream);
    }
}
catch (Exception ex)
{
    Console.WriteLine(ex.ToString());
}
Console.ReadKey();
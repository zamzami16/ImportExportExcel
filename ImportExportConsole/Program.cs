using System.Data;

namespace ImportExportConsole
{
    internal class Program
    {
        public static void Main()
        {
            DataTable dt = AxataPOS.ImportExport.ImportExport.ImportExcel(@"D:\KERJA\AXATA\ImportDataExcel\ImportDataExcel\bin\TestDataTable.xlsx");
        }
    }
}

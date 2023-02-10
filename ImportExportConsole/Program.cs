using System.Data;

namespace ImportExportConsole
{
    internal class Program
    {
        public static void Main()
        {
            DataTable dt = AxataPOS.ImportExport.ImportExport.ImportExcel(@"D:\KERJA\AXATA\ImportExport\test.xlsx");
            AxataPOS.ImportExport.ImportExport.ExportToExcel(dt, @"D:\KERJA\AXATA\ImportExport\test-from-import.xlsx");
        }
    }
}

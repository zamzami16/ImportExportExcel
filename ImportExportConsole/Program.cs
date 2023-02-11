using System;
using System.Data;
using System.Net.Sockets;
using System.Net;
using FirebirdSql.Data.FirebirdClient;

namespace ImportExportConsole
{
    internal class Program
    {
        private static readonly string FileTest = @"D:\KERJA\AXATA\ImportExport\TestImportFromData.xlsx";
        private static readonly string StrCon = $@"database={GetLocalIPAddress()}:C:\Users\yusuf\OneDrive\Desktop\Axata\DB\AxataPOS\EMAS\Debug\Axata.axt;
            user=SYSDBA; Password=masterkey; Connection lifetime=0; Dialect=3; Server=localhost";
        private static readonly FbConnection conn = new FbConnection(StrCon);
        public static void Main()
        {
            /*DataSet ds = AxataPOS.ImportExport.ImportExport.ImportExcel2DataSet(FileTest);
            foreach (DataTable dataTable in ds.Tables)
            {
                Console.WriteLine(dataTable.TableName);
            }
            AxataPOS.ImportExport.ImportExport.ExportToExcel(ds, @"D:\KERJA\AXATA\ImportExport\test-from-import-ds.xlsx");
            
            DataTable dt = AxataPOS.ImportExport.ImportExport.ImportExcel(@"D:\KERJA\AXATA\ImportExport\test.xlsx");
            AxataPOS.ImportExport.ImportExport.ExportToExcel(dt, @"D:\KERJA\AXATA\ImportExport\test-from-import.xlsx");*/

            var watch = System.Diagnostics.Stopwatch.StartNew();
            

            DataTable dt = new DataTable();
            using (var con = conn)
            {
                using (var cmd = new FbCommand())
                {
                    cmd.CommandType = CommandType.Text;
                    cmd.Connection= con;
                    cmd.CommandText = "select * from barang";
                    using (var da = new FbDataAdapter(cmd))
                    {
                        da.Fill(dt);
                    }
                }
            }
            watch.Stop();
            var tread = watch.ElapsedMilliseconds;
            watch.Restart();
            AxataPOS.ImportExport.ImportExport.ExportToExcel(dt, FileTest);

            watch.Stop();
            var twrite = watch.ElapsedMilliseconds;

            Console.WriteLine($"Elapsed reading data {tread} ms.\nElapsed writing excel file(s) {twrite} ms.");
            Console.ReadLine();
        }

        private static string GetLocalIPAddress()
        {
            var host = Dns.GetHostEntry(Dns.GetHostName());
            foreach (var ip in host.AddressList)
            {
                if (ip.AddressFamily == AddressFamily.InterNetwork)
                {
                    return ip.ToString();
                }
            }
            throw new Exception("No network adapters with an IPv4 address in the system!");
        }
    }
}

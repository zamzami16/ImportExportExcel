using AxataPOS.ImportExport.Utils;
using ExcelDataReader;
using System;
using System.Data;
using System.IO;

namespace AxataPOS.ImportExport
{
    public static class ImportExport
    {
        /// <summary>
        /// Import Data From Excel to DataTable
        /// </summary>
        /// <param name="FileName">Nama file excel</param>
        /// <returns>DataTable From Excel</returns>
        public static DataTable ImportExcel(string FileName)
        {
            IExcelDataReader rdr;
            try
            {
                string ext = Path.GetExtension(FileName);
                var ms = File.Open(FileName, FileMode.Open, FileAccess.Read);
                switch (ext)
                {
                    case ".xls": //Old style (2003 or less)
                        rdr = ExcelReaderFactory.CreateBinaryReader(ms);
                        break;
                    case ".xlsx": //New style (2007 or higher)
                        rdr = ExcelReaderFactory.CreateOpenXmlReader(ms);
                        break;
                    default:
                        throw new ArgumentOutOfRangeException(String.Format("File extension {0} is not recognized as valid", ext));
                }

                using (DataSet ds = rdr.AsDataSet(new ExcelDataSetConfiguration
                {
                    UseColumnDataType = true,
                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = true
                    }
                }))
                {
                    DataTable dt = ds.Tables[0];
                    return dt;
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Export DataTable to excel file. Currently support *.xlsx file.
        /// </summary>
        /// <param name="data">DataTable</param>
        /// <param name="FilePath">Nama dan Path dari file</param>
        public static void ExportToExcel(DataTable data, string FilePath)
        {
            try
            {
                var excel = data.ToExcel();
                File.WriteAllBytes(FilePath, excel);
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}

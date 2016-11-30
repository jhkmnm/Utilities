using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Utilities.ExcelHelper
{
    public class ExcelHelper
    {
        public static DataTable GetExcelFile()
        {
            var version = GetExcelVersion() == ExportDocumentType.Excel ? ".xls" : ".xlsx";
            var file = new OpenFileDialog { Title = @"导入Excel", DefaultExt = version, Filter = @"Excel文件|*.xlsx;*.xls" };

            return file.ShowDialog() == DialogResult.OK ? GetExcelFile(file.FileName) : null;
        }

        public static DataTable GetExcelFile(string fileName)
        {
            var sqlconn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1';";            
            var extensionName = GetExcelVersion() == ExportDocumentType.Excel ? ".xls" : ".xlsx";
            if (extensionName == ".xlsx")
            {
                sqlconn = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1';";
            }

            const string sql = "SELECT * FROM [" + "sheet1" + "$]";
            var dsTemp = new DataSet();

            if (System.IO.File.Exists(fileName))
            {
                var oldcom = new System.Data.OleDb.OleDbCommand(sql, new System.Data.OleDb.OleDbConnection(sqlconn));
                var oleda = new System.Data.OleDb.OleDbDataAdapter(oldcom);
                oleda.Fill(dsTemp, "[" + "sheet1" + "$]");

                var dTable = dsTemp.Tables[0].Copy();
                dTable.Clear();
                foreach (DataRow row in dsTemp.Tables[0].Rows)
                {
                    var columns = dsTemp.Tables[0].Columns.Count > 3 ? 3 : dsTemp.Tables[0].Columns.Count;
                    var isnull = true;

                    for (var i = 0; i < columns; i++)
                    {
                        if (row[i].ToString().Length <= 0) continue;
                        isnull = false;
                        break;
                    }

                    if (isnull) break;

                    dTable.ImportRow(row);
                }

                return dTable.Copy();
            }
            return null;
        }

        /// <summary>
        /// 缓存版本号
        /// </summary>
        private static string excelVersion;

        public static ExportDocumentType GetExcelVersion()
        {
            if (excelVersion == null)
            {
                var ap = new Microsoft.Office.Interop.Excel.Application();
                excelVersion = ap.Version;
                QuitExcel(ap);
            }
            float f = excelVersion.ToFloat();
            if (f == 0f)
            {
                return ExportDocumentType.None;
            }
            else if (f < 12f)
            {
                return ExportDocumentType.Excel;
            }
            else
            {
                return ExportDocumentType.Excel2007;
            }
        }

        private static void QuitExcel(Microsoft.Office.Interop.Excel.Application application)
        {
            application.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(application);
        }

        public enum ExportDocumentType
        {
            None,
            Excel,
            CSV,
            PDF,
            Excel2007
        }
    }
}

using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Data.OleDb;
using System.Diagnostics;
using System.Xml.Serialization;
using System.Text.RegularExpressions;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;

namespace Utilities
{
    public class ExcelHelper : IDisposable
    {
        #region Fields
        private static ExcelHelper instance;
        private static readonly object syncRoot = new object();
        private string returnMessage;
        private Excel.Application xlApp;
        private Excel.Workbooks workbooks = null;
        private Excel.Workbook workbook = null;
        private Excel.Worksheet worksheet = null;
        private Excel.Range range = null;
        private int status = -1;
        private bool disposed = false;//是否已经释放资源的标记
        #endregion

        private ExcelHelper()
        {
            status = IsExistExecl() ? 0 : -1;
        }

        public static ExcelHelper GetInstance()
        {
            return new ExcelHelper();
        }

        #region Properties
        /// <summary>
        /// 返回信息
        /// </summary>
        public string ReturnMessage
        {
            get { return returnMessage; }
        }

        /// <summary>
        /// 状态:0-正常，-1-失败 1-成功
        /// </summary>
        public int Status
        {
            get { return status; }
        }
        #endregion

        #region Methods
        /// <summary>
        /// 判断是否安装Excel
        /// </summary>
        /// <returns></returns>
        protected bool IsExistExecl()
        {
            try
            {
                xlApp = new Excel.Application();
                if (xlApp == null)
                {
                    returnMessage = "无法创建Excel对象，可能您的计算机未安装Excel!";
                    return false;
                }
            }
            catch (Exception ex)
            {
                returnMessage = "请正确安装Excel！";
                return false;
            }

            return true;
        }

        /// <summary>
        /// 获得保存路径
        /// </summary>
        /// <returns></returns>
        public static string SaveFileDialog()
        {
            var version = GetExcelVersion() == ExportDocumentType.Excel ? ".xls" : ".xlsx";

            SaveFileDialog sfd = new SaveFileDialog();
            sfd.DefaultExt = version;
            sfd.Filter = @"Excel文件|*.xlsx;*.xls";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                return sfd.FileName;
            }
            return string.Empty;
        }

        /// <summary>
        /// 获得打开文件的路径
        /// </summary>
        /// <returns></returns>
        public static string OpenFileDialog()
        {
            var version = GetExcelVersion() == ExportDocumentType.Excel ? ".xls" : ".xlsx";

            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = version;
            ofd.Filter = @"Excel文件|*.xlsx;*.xls";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                return ofd.FileName;
            }
            return string.Empty;
        }

        /// <summary>
        /// 设置单元格边框
        /// </summary>
        protected void SetCellsBorderAround()
        {
            range.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, null);

            range.Borders[Excel.XlBordersIndex.xlInsideVertical].ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
            range.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlThin;
        }

        public bool DataTableToExcel(DataTable dt, string sheetName)
        {
            string filename = SaveFileDialog();
            return DataTableToExecl(dt, filename, sheetName);
        }

        /// <summary>
        /// 将DataTable导出Excel
        /// </summary>
        /// <param name="dt">数据集</param>
        /// <param name="saveFilePath">保存路径</param>
        /// <param name="reportName">报表名称</param>
        /// <returns>是否成功</returns>
        public bool DataTableToExecl(DataTable dt, string saveFileName, string reportName)
        {
            //判断是否安装Excel
            bool fileSaved = false;
            if (status == -1) return fileSaved;
            //判断数据集是否为null
            if (dt == null)
            {
                returnMessage = "无引出数据！";
                return false;
            }
            //判断保存路径是否有效
            if (!saveFileName.Contains(":"))
            {
                returnMessage = "引出路径有误！请选择正确路径！";
                return false;
            }

            //创建excel对象
            workbooks = xlApp.Workbooks;
            workbook = workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            worksheet = (Excel.Worksheet)workbook.Worksheets[1];//取得sheet1
            worksheet.Cells.Font.Size = 10;
            worksheet.Cells.NumberFormat = "@";
            long totalCount = dt.Rows.Count;
            long rowRead = 0;
            float percent = 0;
            int rowIndex = 0;

            //第一行为报表名称，如果为null则不保存该行    
            ++rowIndex;
            worksheet.Cells[rowIndex, 1] = reportName;
            range = (Excel.Range)worksheet.Cells[rowIndex, 1];
            range.Font.Bold = true;

            //写入字段(标题)
            ++rowIndex;
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                worksheet.Cells[rowIndex, i + 1] = dt.Columns[i].ColumnName;
                range = (Excel.Range)worksheet.Cells[rowIndex, i + 1];

                range.Font.Color = ColorTranslator.ToOle(Color.Blue);
                range.Interior.Color = dt.Columns[i].Caption == "表体" ? ColorTranslator.ToOle(Color.SkyBlue) : ColorTranslator.ToOle(Color.Yellow);
            }

            //写入数据
            ++rowIndex;
            for (int r = 0; r < dt.Rows.Count; r++)
            {
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    worksheet.Cells[r + rowIndex, i + 1] = dt.Rows[r][i].ToString();
                }
                rowRead++;
                percent = ((float)(100 * rowRead)) / totalCount;
            }

            //画单元格边框
            range = worksheet.get_Range(worksheet.Cells[2, 1], worksheet.Cells[dt.Rows.Count + 2, dt.Columns.Count]);
            this.SetCellsBorderAround();

            //列宽自适应
            range.EntireColumn.AutoFit();

            //保存文件
            if (saveFileName != "")
            {
                try
                {
                    workbook.Saved = true;
                    workbook.SaveCopyAs(saveFileName);
                    fileSaved = true;
                }
                catch (Exception ex)
                {
                    fileSaved = false;
                    returnMessage = "导出文件时出错,文件可能正被打开！\n" + ex.Message;
                }
            }
            else
            {
                fileSaved = false;
            }

            //释放Excel对应的对象（除xlApp,因为创建xlApp很花时间，所以等析构时才删除)
            //Dispose(false);
            Dispose();
            return fileSaved;
        }

        public DataSet ImportExcel()
        {
            return ImportExcel(OpenFileDialog());
        }

        /// <summary>
        /// 导入EXCEL到DataSet
        /// </summary>
        /// <param name="fileName">Excel全路径文件名</param>
        /// <returns>导入成功的DataSet</returns>
        public DataSet ImportExcel(string fileName)
        {
            if (status == -1) return null;
            //判断文件是否被其他进程使用            
            try
            {
                workbook = xlApp.Workbooks.Open(fileName, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, 1, 0);
                worksheet = (Excel.Worksheet)workbook.Worksheets[1];
            }
            catch
            {
                returnMessage = "Excel文件处于打开状态，请保存关闭";
                return null;
            }

            //获得所有Sheet名称
            int n = workbook.Worksheets.Count;
            string[] sheetSet = new string[n];
            for (int i = 0; i < n; i++)
            {
                sheetSet[i] = ((Excel.Worksheet)workbook.Worksheets[i + 1]).Name;
            }

            //释放Excel相关对象
            Dispose();

            //把EXCEL导入到DataSet
            DataSet ds = null;            
            List<string> connStrs = new List<string>();
            connStrs.Add("Provider = Microsoft.Jet.OLEDB.4.0; Data Source = " + fileName + ";Extended Properties=\"Excel 8.0;HDR=No;IMEX=1;\"");
            connStrs.Add("Provider = Microsoft.ACE.OLEDB.12.0 ; Data Source = " + fileName + ";Extended Properties=\"Excel 12.0;HDR=No;IMEX=1;\"");
            foreach (string connStr in connStrs)
            {
                ds = GetDataSet(connStr, sheetSet);
                if (ds != null) break;
            }
            return ds;
        }

        /// <summary>
        /// 通过olddb获得dataset
        /// </summary>
        /// <param name="connectionstring"></param>
        /// <returns></returns>
        protected DataSet GetDataSet(string connStr, string[] sheetSet)
        {
            DataSet ds = null;
            using (OleDbConnection conn = new OleDbConnection(connStr))
            {
                try
                {
                    conn.Open();
                    OleDbDataAdapter da;
                    ds = new DataSet();
                    for (int i = 0; i < sheetSet.Length; i++)
                    {
                        string sql = "select * from [" + sheetSet[i] + "$] ";
                        da = new OleDbDataAdapter(sql, conn);
                        da.Fill(ds, sheetSet[i]);
                        da.Dispose();
                    }
                    conn.Close();
                    conn.Dispose();
                }
                catch (Exception ex)
                {
                    return null;
                }
            }
            return ds;
        }

        /// <summary>
        /// 释放Excel对应的对象资源
        /// </summary>
        /// <param name="isDisposeAll"></param>
        protected virtual void Dispose(bool disposing)
        {
            try
            {
                if (!disposed)
                {
                    if (disposing)
                    {
                        if (range != null)
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                            range = null;
                        }
                        if (worksheet != null)
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                            worksheet = null;
                        }
                        if (workbook != null)
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                            workbook = null;
                        }
                        if (workbooks != null)
                        {
                            xlApp.Application.Workbooks.Close();
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbooks);
                            workbooks = null;
                        }
                        if (xlApp != null)
                        {
                            xlApp.Quit();
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
                        }
                        int generation = GC.GetGeneration(xlApp);
                        System.GC.Collect(generation);
                    }

                    //非托管资源的释放
                    //KillExcel();
                }
                disposed = true;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary> 
        /// 会自动释放非托管的该类实例的相关资源
        /// </summary>
        public void Dispose()
        {
            try
            {
                Dispose(true);
                //告诉垃圾回收器,资源已经被回收
                GC.SuppressFinalize(this);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 关闭
        /// </summary>
        public void Close()
        {
            try
            {
                this.Dispose();
            }
            catch (Exception e)
            {

                throw e;
            }
        }

        /// <summary>
        /// 析构函数
        /// </summary>
        ~ExcelHelper()
        {
            try
            {
                Dispose(false);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 关闭Execl进程(非托管资源使用)
        /// </summary>
        private void KillExcel()
        {
            try
            {
                Process[] ps = Process.GetProcesses();
                foreach (Process p in ps)
                {
                    if (p.ProcessName.ToLower().Equals("excel"))
                    {
                            p.Kill();                     
                    }
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("ERROR " + ex.Message);
            }
        }

        private static System.Xml.XmlNode GetXmlNode(string dataKey)
        {
            XmlHelper.LoadXml("Config/ExcelModel.xml");
            return XmlHelper.GetXmlNode("ViewDataDetail", dataKey);
        }

        /// <summary>
        /// 读取Excel并将数据返回
        /// 如果存在错误数据，将返回一个count=0的List
        /// </summary>
        /// <returns></returns>
        public List<ImportXMLModel> Import(string dataKey, ref List<ErrorVO> importError)
        {
            var imports = new List<ImportXMLModel>();
            var dv = new DataVerifier();

            var _quotationNode = GetXmlNode(dataKey);
            var importmodel = XmlHelper.XmlDeserialize<ImportXMLModel>(_quotationNode.OuterXml, Encoding.UTF8);

            var dt = ImportExcel().Tables[0];

            dv.Check(dt == null, "导入的Excel中没有数据");

            if (dv.Pass)
            {
                var checkcol = importmodel.Columns.Where(item => !dt.Columns.Contains(item.ColumnName));
                dv.CheckIfBeforePass(checkcol.Any(), "导入的格式与模板不一致,请按模板导入");

                if (dv.Pass)
                {
                    for (var i = 0; i < dt.Rows.Count; i++)
                    {
                        var import = Clone(importmodel);
                        foreach (var col in importmodel.Columns)
                        {
                            import[col.ColumnName].ColumnVaue = dt.Rows[i][col.ColumnName].ToString().Trim();
                        }
                        import.Columns.Add(new ImportColumn { ColumnName = "序号", ColumnVaue = (i + 2).ToString() });

                        if (!import.Pass)
                        {
                            var error = new ErrorVO("第 " + (i + 2).ToString() + " 行", import.ErrorMessage);
                            importError.Add(error);
                        }
                        else
                        {
                            imports.Add(import);
                        }
                    }
                }
            }
            dv.ShowMsgIfFailed();

            return imports;
        }

        /// <summary>
        /// 利用 System.Runtime.Serialization序列化与反序列化完成引用对象的复制
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="realObject"></param>
        /// <returns></returns>
        public static T Clone<T>(T realObject)
        {
            using (Stream objectStream = new MemoryStream())
            {
                IFormatter formatter = new BinaryFormatter();
                formatter.Serialize(objectStream, realObject);
                objectStream.Seek(0, SeekOrigin.Begin);
                return (T)formatter.Deserialize(objectStream);
            }
        }

        #endregion

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

    [Serializable]
    [XmlType("ViewDataDetail", Namespace = "")]
    public class ImportXMLModel
    {
        [XmlAttribute]
        public string FileName
        { get; set; }

        [XmlAttribute]
        public string DataKey
        { get; set; }

        [XmlElement("DataColumn")]
        public List<ImportColumn> Columns
        { get; set; }

        public ImportColumn this[string columnName]
        {
            get
            {
                return Columns.Single(where => where.ColumnName == columnName);
            }
        }

        public bool Pass
        {
            get
            {
                var v = Columns.Where(where => !where.Pass);
                if (v.Count() > 0)
                {
                    _errormessage = v.ToList()[0].ErrorInfo;
                    return false;
                }

                return true;
            }
        }

        private string _errormessage;
        public string ErrorMessage
        {
            get
            {
                return _errormessage;
            }
        }
    }

    [Serializable]
    [XmlType("DataColumn")]
    public class ImportColumn
    {
        [XmlAttribute]
        public string ColumnName
        { get; set; }

        [XmlAttribute]
        public string RegStr
        { get; set; }

        [XmlAttribute]
        public int MaxLength
        { get; set; }

        [XmlAttribute]
        public string DataType
        { get; set; }

        string _errorInfo;
        [XmlAttribute]
        public string ErrorInfo
        { get { return "【" + ColumnName + "】" + _errorInfo; } set { _errorInfo = value; } }

        [XmlAttribute]
        public bool IsNull
        { get; set; }

        private string _columnvalue = "";
        [XmlElement("")]
        public string ColumnVaue
        {
            get { return _columnvalue; }
            set
            {
                if (!IsNull && string.IsNullOrEmpty(value))
                {
                    _pass = false;
                }
                else if (MaxLength > 0 && value.Length > MaxLength)
                {
                    _pass = false;
                    ErrorInfo = " 超出允许的长度，最大长度为 " + MaxLength.ToString();
                }
                else if (!string.IsNullOrEmpty(RegStr) && !string.IsNullOrEmpty(value) && !string.IsNullOrEmpty(value))
                {
                    var r = new Regex(RegStr);
                    var m = r.Match(value);
                    _pass = m.Success;
                }
                else if (!string.IsNullOrEmpty(DataType) && !string.IsNullOrEmpty(value) && !CheckDataType(DataType, value))
                {
                    _pass = false;
                }
                _columnvalue = value.Replace(' ', ' ');
            }
        }

        private bool _pass = true;
        public bool Pass
        {
            get { return _pass; }
        }

        private bool CheckDataType(string dataType, string value)
        {
            var bresult = true;

            switch (dataType.ToUpper())
            {
                case "DATETIME":
                    DateTime result;
                    bresult = DateTime.TryParse(value, out result);
                    break;
            }

            return bresult;
        }
    }

    public class ErrorVO
    {
        private string _rowindex;
        private string _message;

        public ErrorVO(string rowindex, string message)
        {
            _rowindex = rowindex;
            _message = message;
        }

        public string 序号
        { get { return _rowindex; } }

        public string 错误信息
        { get { return _message; } }
    }
}
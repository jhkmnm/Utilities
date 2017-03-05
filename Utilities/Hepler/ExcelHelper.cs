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
    /// <summary>
    /// 导出Excel辅助类
    /// 郭龙飞 2013-5-3
    /// </summary>
    public static class ExportExcelHelper
    {
        #region 导出对象数据到Excel文件
        /// <summary>
        /// 高效的导出方法，
        /// 如果记录太多可以分sheet页了
        /// </summary>
        /// <param name="gridView">DataGridView控件</param>
        /// <param name="fileName">文件名</param>
        /// <param name="sheetName">sheet名</param>
        public static bool ExportExcel(DataGridView gridView, string fileName, string sheetName)
        {
            return ExportExcel(gridView, fileName, sheetName, null);
        }

        /// <summary>
        /// 高效的导出方法，
        /// 如果记录太多可以分sheet页了
        /// </summary>
        /// <param name="gridView">DataGridView控件</param>
        /// <param name="fileName">文件名</param>
        /// <param name="sheetName">sheet名</param>
        /// <param name="nonColumns">不显示列</param>
        public static bool ExportExcel(DataGridView gridView, string fileName, string sheetName, List<string> nonColumns)
        {
            string version = GetExcelExtension();
            if (string.IsNullOrEmpty(version))
            {
                return false;
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "导出Excel";
            saveFileDialog.DefaultExt = version;
            saveFileDialog.Filter = "Excel文件|*" + version;
            saveFileDialog.FileName = fileName;
            if (saveFileDialog.ShowDialog() != DialogResult.OK)
            {
                return false;
            }
            try
            {
                bool bResult = SaveExcelFile(gridView, saveFileDialog.FileName, version, sheetName, nonColumns);
                if (bResult == true)
                {
                    return true;
                }
            }
            catch
            {
            }
            return false;
        }

        /// <summary>
        /// 童荣辉增加 20130724 抽取出共用的代码，以适应与弹出保存框区分开
        /// DataGridView数据展出到Excel
        /// </summary>
        private static bool SaveExcelFile(DataGridView gridView, string strfileName, string strVersion, string sheetName, List<string> nonColumns)
        {
            int maxLength = 65535;
            if (strVersion.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                maxLength = 1048575;
            }
            string saveFileName = strfileName;
            System.Reflection.Missing miss = System.Reflection.Missing.Value;
            //创建EXCEL对象appExcel,Workbook对象,Worksheet对象,Range对象
            ExcelApp.Application appExcel = new ExcelApp.Application();
            ExcelApp.Workbook workbookData = appExcel.Workbooks.Add(ExcelApp.XlWBATemplate.xlWBATWorksheet);
            ExcelApp.Worksheet worksheetData = null;
            ExcelApp.Range rangedata;
            //设置对象不可见
            appExcel.Visible = false;
            int countOfSheets = ValueCalculate.CalculateCountOfPage(gridView.RowCount, maxLength);
            /* 在调用Excel应用程序，或创建Excel工作簿之前，记着加上下面的两行代码
             * 这是因为Excel有一个Bug，如果你的操作系统的环境不是英文的，而Excel就会在执行下面的代码时，报异常。
             */
            for (int ipage = 1; ipage <= countOfSheets; ipage++)
            {
                if (worksheetData == null)
                {
                    worksheetData = (Microsoft.Office.Interop.Excel.Worksheet)workbookData.Worksheets.Add(Type.Missing, Type.Missing, 1, Type.Missing);
                }
                else
                {
                    worksheetData = (Microsoft.Office.Interop.Excel.Worksheet)workbookData.Worksheets.Add(Type.Missing, worksheetData, 1, Type.Missing);
                }
                //当前页数据条数
                int currentPageNums = 0;
                //给工作表赋名称
                if (countOfSheets == 1)
                {
                    worksheetData.Name = sheetName;
                    currentPageNums = gridView.RowCount;
                }
                else
                {
                    worksheetData.Name = string.Format("{0}{1}", sheetName, ipage);
                    if (ipage == countOfSheets)
                    {
                        currentPageNums = gridView.RowCount - maxLength * (ipage - 1);
                    }
                    else
                    {
                        currentPageNums = maxLength;
                    }
                }
                //新建一个字典来控制Visible的列不导出
                Dictionary<int, int> dictionary = new Dictionary<int, int>();
                int iVisible = 0;
                //清零计数并开始计数
                // 保存到WorkSheet的表头，你应该看到，是一个Cell一个Cell的存储，这样效率特别低，解决的办法是，使用Rang，一块一块地存储到Excel
                for (int i = 0; i < gridView.ColumnCount; i++)
                {
                    if (gridView.Columns[i].Visible &&
                        ((nonColumns != null && !nonColumns.Contains(gridView.Columns[i].Name)) || nonColumns == null))
                    {
                        worksheetData.Cells[1, iVisible + 1] = gridView.Columns[i].HeaderText.ToString();
                        var range = worksheetData.Cells[1, iVisible + 1];
                        range.Font.Bold = true;
                        range.Font.Size = 12;
                        dictionary.Add(iVisible, i);
                        iVisible++;
                    }
                }
                //先给Range对象一个范围为A2开始，Range对象可以给一个CELL的范围，也可以给例如A1到H10这样的范围
                //因为第一行已经写了表头，所以所有数据都应该从A2开始
                rangedata = worksheetData.get_Range("A2", miss);
                Microsoft.Office.Interop.Excel.Range xlRang = null;

                //iColumnAccount为实际列数，最大列数
                int iColumnAccount = iVisible;
                //在内存中声明一个iEachSize×iColumnAccount的数组，iEachSize是每次最大存储的行数，iColumnAccount就是存储的实际列数
                //object[,] objVal = new object[currentPageNums, iColumnAccount];
                //每次最大导入数据量
                int perMaxCount = 2000;
                //当前行
                int iParstedRow = 0;
                int times = ValueCalculate.CalculateCountOfPage(currentPageNums, perMaxCount);
                for (int ti = 0; ti < times; ti++)
                {
                    //当前循环的得到的数
                    int currentTimeNum = 0;
                    if ((currentPageNums - ti * perMaxCount) < perMaxCount)
                    {
                        currentTimeNum = currentPageNums - ti * perMaxCount;
                    }
                    else
                    {
                        currentTimeNum = perMaxCount;
                    }
                    object[,] objVal = new object[currentTimeNum, iColumnAccount];
                    for (int i = 0; i < currentTimeNum; i++)
                    {
                        for (int j = 0; j < iColumnAccount; j++)
                        {
                            int numOfColumn;
                            if (!dictionary.TryGetValue(j, out numOfColumn))
                            {
                                throw new KeyNotFoundException("导出Excel异常，未找到列！");
                            }
                            object cellValue = gridView[numOfColumn, ti * perMaxCount + i + (ipage - 1) * maxLength].Value;
                            if (cellValue != null && (cellValue.GetType() == typeof(string) || cellValue.GetType() == typeof(DateTime)))
                            {
                                cellValue = "'" + cellValue.ToString();
                            }
                            objVal[i, j] = cellValue;
                        }
                        System.Windows.Forms.Application.DoEvents();
                    }
                    xlRang = worksheetData.get_Range("A" + (iParstedRow + 2).ToString(), (ExcelColumnNameEnum.A + iColumnAccount - 1).ToString() + (currentTimeNum + iParstedRow + 1).ToString());
                    // 调用Range的Value2属性，把内存中的值赋给Excel
                    xlRang.Value2 = objVal;

                    iParstedRow += currentTimeNum;
                }
            }
            try
            {
                workbookData.Saved = true;
                workbookData.SaveCopyAs(saveFileName);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
            finally
            {
                QuitExcel(appExcel);
            }
        }
        #endregion

        /// <summary>
        /// 退出Excel
        /// </summary>
        /// <param name="application">Excel进程</param>
        private static void QuitExcel(Microsoft.Office.Interop.Excel.Application application)
        {
            application.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(application);
        }

        /// <summary>
        /// Excel扩展名
        /// </summary>
        private static string excelVersion;

        /// <summary>
        /// 获取Excel扩展名
        /// </summary>
        /// <returns>Excel扩展名</returns>
        private static string GetExcelExtension()
        {
            if (excelVersion == null)
            {
                Microsoft.Office.Interop.Excel.Application ap = new Microsoft.Office.Interop.Excel.Application();
                excelVersion = ap.Version;
                QuitExcel(ap);
            }
            float f = excelVersion.ToFloat();
            if (f == 0f)
            {
                MessageBox.Show("请安装EXCEL 2003或2007");
                return "";
            }
            else if (f < 12f)
            {
                return ".xls";
            }
            else
            {
                return ".xlsx";
            }
        }

        public static void OpenFile(string fileName)
        {
            if (UTIL.MsgTool.ShowConfirmMessage("需要打开吗？") == DialogResult.OK)
            {
                try
                {
                    System.Diagnostics.Process.Start(fileName);
                }
                catch (Exception)
                {
                    UTIL.MsgTool.ShowMessage("系统中没有能打开" + fileName + "的程序");
                }
            }
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
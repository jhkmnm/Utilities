using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Utilities
{
    public class InitComboBox
    {
        public static void InitDorpDownByEnum(ComboBox combobox, Type enumType)
        {
            InitDorpDownByEnum(combobox, enumType, false, "");
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="combobox"></param>
        /// <param name="enumType"></param>
        /// <param name="defaultvalue">默认会加上全部，全部的默认值可以指定</param>
        /// <param name="values">可能只需要显示部分值，可以在这个参数中指定</param>
        public static void InitDorpDownByEnum(ComboBox combobox, Type enumType, string defaultvalue, string[] values = null)
        {
            InitDorpDownByEnum(combobox, enumType, true, defaultvalue, values);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="combobox"></param>
        /// <param name="enumType"></param>
        /// <param name="addFirstItem">是否增加全部这一项</param>
        /// <param name="defaultvalue">默认会加上全部，全部的默认值可以指定</param>
        /// <param name="values">可能只需要显示部分值，可以在这个参数中指定</param>
        public static void InitDorpDownByEnum(ComboBox combobox, Type enumType, bool addFirstItem, string defaultvalue, string[] values = null)
        {
            EnumDescConverter enumdescconverter = new EnumDescConverter(enumType);
            string[] names = Enum.GetNames(enumType);
            DataTable dtSource = new DataTable();
            dtSource.Columns.Add("Text");
            dtSource.Columns.Add("Value");

            if (addFirstItem)
            {
                DataRow firstrow = dtSource.NewRow();
                firstrow["Value"] = defaultvalue;
                firstrow["Text"] = "全部";
                dtSource.Rows.Add(firstrow);
            }

            for (int i = 0; i < names.Length; i++)
            {
                if (values == null || (values != null && values.Contains(names[i])))
                {
                    DataRow row = dtSource.NewRow();
                    row["Value"] = names[i];
                    row["Text"] = (string)enumdescconverter.ConvertTo(names[i], typeof(string));
                    dtSource.Rows.Add(row);
                }
            }

            combobox.DataSource = dtSource;
            combobox.DisplayMember = "Text";
            combobox.ValueMember = "Value";
        }
    }
}

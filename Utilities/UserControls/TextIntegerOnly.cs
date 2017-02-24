using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Utilities.UserControls
{
    public partial class TextIntegerOnly : UserControl
    {
        public TextIntegerOnly()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 允许负数
        /// </summary>
        public bool IsNegativeNumbers { get; set; }

        /// <summary>
        /// 允许小数
        /// </summary>
        public bool IsDecimal { get; set; }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if(e.KeyChar == (char)45)
            {
                if(!IsNegativeNumbers || !string.IsNullOrEmpty(textBox1.Text))
                {
                    e.Handled = true;
                }
            }
            else if(e.KeyChar == (char)46)
            {
                if (!IsDecimal || string.IsNullOrEmpty(textBox1.Text) || textBox1.Text == "-" || textBox1.Text.Contains("."))
                {
                    e.Handled = true;
                }
            }
            else if(!char.IsNumber(e.KeyChar) && e.KeyChar != (char)8)  // 允许输入退格键
            {
                e.Handled = true;   // 经过判断为数字，可以输入
            }
        }

        public new string Text { 
            get { return textBox1.Text; }
            set { textBox1.Text = value; }
        }
    }
}

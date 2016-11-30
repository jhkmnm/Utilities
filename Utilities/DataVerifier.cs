using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Utilities
{
    [DebuggerStepThrough]
    public class DataVerifier
    {
        private readonly StringBuilder content;
        private Type customExceptionType;

        public event Action<bool> Verified;

        public DataVerifier()
        {
            this.content = new StringBuilder();
            this.customExceptionType = null;
            this.Pass = true;
            this.PromptThrowException = false;
        }

        public DataVerifier(Type customExceptionType)
            : this(true, customExceptionType, "", new object[0])
        {
        }

        public DataVerifier(bool promptThrowException, Type customExceptionType)
            : this(promptThrowException, customExceptionType, "", new object[0])
        {
        }

        public DataVerifier(string title, params object[] args)
            : this(false, null, title, args)
        {
        }

        public DataVerifier(bool promptThrowException, Type customExceptionType, string title, params object[] args)
            : this()
        {
            PromptThrowException = promptThrowException;
            this.customExceptionType = customExceptionType;
            Title = string.Format(title, args);
        }

        public void AddContent(string msg, params object[] args)
        {
            Check(true, msg, args);
        }

        public void Check(bool expression, string msg, params object[] args)
        {
            if (Pass)
            {
                Pass = !expression;
            }
            if (!(((msg.EndsWith("！") || msg.EndsWith("!")) || msg.EndsWith("。")) || msg.EndsWith(".")))
            {
                msg = msg + "！";
            }
            this.content.Append(expression ? (string.Format(msg, args) + "\r\n") : "");
            this.RaiseVerified(expression);
            if (this.PromptThrowException)
            {
                this.ThrowExceptionIfFailed();
            }
        }

        public void Check(bool expression, Action beforeAction, string msg, params object[] args)
        {
            beforeAction();
            this.Check(expression, msg, new object[0]);
        }

        public void CheckIfBeforePass(bool expression, string msg, params object[] args)
        {
            if (this.Pass)
            {
                this.Check(expression, msg, args);
            }
        }

        public void CustomShowMsgIfFailed(Action<string> showMsgMethod)
        {
            if (!(this.Pass || (showMsgMethod == null)))
            {
                showMsgMethod(this.GetContent());
            }
        }

        private string GetContent()
        {
            return (this.Title + "\r\n" + this.Content);
        }

        private void RaiseVerified(bool expression)
        {
            if (this.Verified != null)
            {
                this.Verified(expression);
            }
        }

        public void ShowMsgIfFailed()
        {
            if (!this.Pass)
            {
                MessageBox.Show(this.GetContent(), @"提示", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
        }

        public void ThrowExceptionIfFailed()
        {
            if (!this.Pass)
            {
                Exception exception = null;
                if (this.CustomExceptionType == null)
                {
                    exception = new Exception(this.GetContent());
                }
                else
                {
                    exception = Activator.CreateInstance(this.CustomExceptionType, new object[] { this.GetContent() }) as Exception;
                }
                throw exception;
            }
        }

        public string Content
        {
            get
            {
                return this.content.ToString();
            }
            set
            {
                this.content.AppendLine(value);
            }
        }

        public Type CustomExceptionType
        {
            get
            {
                return this.customExceptionType;
            }
            set
            {
                if (!value.IsAssignableFrom(typeof(Exception)))
                {
                    throw new Exception("CustomExceptionType必须继承自Exception");
                }
                this.customExceptionType = value;
            }
        }

        public bool Pass { get; private set; }

        public bool PromptThrowException { get; set; }

        public string Title { get; set; }
    }
}

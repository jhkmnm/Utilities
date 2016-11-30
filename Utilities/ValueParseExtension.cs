using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Utilities
{
    public static class ValueParseExtension
    {
        /// <summary>
        /// 转型成浮点型
        /// 如果转型不成功则返回0
        /// </summary>
        /// <param name="oldValue">string类型的值</param>
        /// <returns>float类型</returns>
        public static float ToFloat(this string oldValue)
        {
            float newValue;
            float.TryParse(oldValue, out newValue);
            return newValue;
        }

        /// <summary>
        /// 转型成浮点型
        /// 如果转型不成功则返回默认值
        /// </summary>
        /// <param name="oldValue">string类型的值</param>
        /// <param name="defaultValue">默认值</param>
        /// <returns>float类型</returns>
        public static float ToFloat(this string oldValue, float defaultValue)
        {
            float newValue;
            if (float.TryParse(oldValue, out newValue))
            {
                return newValue;
            }
            else
            {
                return defaultValue;
            }
        }

        /// <summary>
        /// 转型成Int32类型
        /// 如果转型不成功则返回0
        /// </summary>
        /// <param name="oldValue">string类型的值</param>
        /// <returns>Int32类型</returns>
        public static Int32 ToInt32(this string oldValue)
        {
            Int32 newValue;
            Int32.TryParse(oldValue, out newValue);
            return newValue;
        }

        /// <summary>
        /// 转型成Int32类型
        /// </summary>
        /// <param name="oldValue">string类型的值</param>
        /// <param name="defaultValue">默认值</param>
        /// <returns>Int32类型</returns>
        public static Int32 ToInt32(this string oldValue, Int32 defaultValue)
        {
            Int32 newValue;
            if (Int32.TryParse(oldValue, out newValue))
            {
                return newValue;
            }
            else
            {
                return defaultValue;
            }
        }

        /// <summary>
        /// 转型成bool类型
        /// </summary>
        /// <param name="oldValue">string类型的值</param>
        /// <returns>bool类型</returns>
        public static bool ToBool(this string oldValue)
        {
            bool newValue;
            bool.TryParse(oldValue, out newValue);
            return newValue;
        }

        /// <summary>
        /// 转型成bool类型
        /// </summary>
        /// <param name="oldValue">string类型的值</param>
        /// <param name="defaultValue">默认值</param>
        /// <returns>bool类型</returns>
        public static bool ToBool(this string oldValue, bool defaultValue)
        {
            bool newValue;
            if (bool.TryParse(oldValue, out newValue))
            {
                return newValue;
            }
            else
            {
                return defaultValue;
            }
        }

        /// <summary>
        /// 转型成double类型
        /// </summary>
        /// <param name="oldValue">string类型的值</param>
        /// <returns>double类型</returns>
        public static double ToDoubleNew(this string oldValue)
        {
            double newValue;
            double.TryParse(oldValue, out newValue);
            return newValue;
        }

        /// <summary>
        /// 转型成double类型
        /// </summary>
        /// <param name="oldValue">string类型的值</param>
        /// <param name="defaultValue">默认值</param>
        /// <returns>double类型</returns>
        public static double ToDoubleNew(this string oldValue, double defaultValue)
        {
            double newValue;
            if (double.TryParse(oldValue, out newValue))
            {
                return newValue;
            }
            else
            {
                return defaultValue;
            }
        }

        /// <summary>
        /// 泛型转型，转型不成功将抛出FormatException异常
        /// </summary>
        /// <typeparam name="T">期待的类型</typeparam>
        /// <param name="oldValue">string类型的值</param>
        /// <returns>期待的类型</returns>
        public static T ToGeneric<T>(this string oldValue)
        {
            try
            {
                T newValue = (T)Convert.ChangeType(oldValue, typeof(T));
                return newValue;
            }
            catch (FormatException)
            {
                return default(T);
            }
        }

        /// <summary>
        /// 泛型转型，转型不成功将抛出FormatException异常
        /// </summary>
        /// <typeparam name="T">期待的类型</typeparam>
        /// <param name="oldValue">string类型的值</param>
        /// <returns>期待的类型</returns>
        public static T ToGeneric<T>(this string oldValue, T defaultValue)
        {
            try
            {
                T newValue = (T)Convert.ChangeType(oldValue, typeof(T));
                return newValue;
            }
            catch (FormatException)
            {
                return defaultValue;
            }
        }

        /// <summary>
        /// 转型成时间类型
        /// </summary>
        /// <param name="oldValue">string类型的值</param>
        /// <returns>DateTime类型</returns>
        public static DateTime ToDateTime(this string oldValue)
        {
            DateTime newValue;
            DateTime.TryParse(oldValue, out newValue);
            return newValue;
        }

        /// <summary>
        /// 转型成时间类型
        /// </summary>
        /// <param name="oldValue">string类型的值</param>
        /// <param name="defaultValue">默认值</param>
        /// <returns>DateTime类型</returns>
        public static DateTime ToDateTime(this string oldValue, DateTime defaultValue)
        {
            DateTime newValue;
            if (DateTime.TryParse(oldValue, out newValue))
            {
                return newValue;
            }
            else
            {
                return defaultValue;
            }
        }
    }
}

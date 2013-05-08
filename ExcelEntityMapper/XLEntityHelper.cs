using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;

namespace ExcelEntityMapper
{
    /// <summary>
    /// A static class which helps for converting values from / to property types into excel values and viceversa.
    /// </summary>
    public static class XLEntityHelper
    {
        private static readonly string[] DateFormats = { "dd/MM/yyyy", "dd/MM/yyyy hh:mm:ss" };
        private static readonly HashSet<Type> DateTypes = new HashSet<Type>();

        /// <summary>
        /// 
        /// </summary>
        static XLEntityHelper()
        {
            DateTypes.Add(typeof(DateTime));
            DateTypes.Add(typeof(DateTime?));
        }

        /// <summary>
        /// Indicates if the argument is considered an Datetime type
        /// </summary>
        /// <param name="type">The type to evaluate.</param>
        /// <returns>returns true if the argument type is an DateTime or DateTime?</returns>
        public static bool IsDateTime(Type type)
        {
            return DateTypes.Contains(type);
        }

        /// <summary>
        /// Gets a nullable DateTime object if the string format match with the given formats.
        /// </summary>
        /// <param name="value">The datetime string value to convert.</param>
        /// <param name="formats">The formats to consider for transforming the given string value.</param>
        /// <returns></returns>
        public static DateTime? GetNullableDateTime(string value, string[] formats)
        {
            DateTime result;

            if (formats == null)
                throw new NullReferenceException("Formats cannot be null");

            if (formats.Length == 0)
                formats = DateFormats;

            bool esito = DateTime.TryParseExact(value, formats, DateTimeFormatInfo.CurrentInfo, DateTimeStyles.AssumeLocal, out result);
            if (esito)
                return result;

            return (DateTime?)null;
        }

        /// <summary>
        /// Normalize the argument, so modifies the string value if it has empty chars at the end if the given argument is not null.
        /// </summary>
        /// <param name="value">The string to normalize.</param>
        /// <returns>returns null if the given argument is null or It contains only empty chars.</returns>
        public static string NormalizeValue(string value)
        {
            if (value != null)
            {
                value = value.Trim();
                return value == string.Empty ? null : value;
            }
            return null;
        }

        /// <summary>
        /// Converts the given string into the given generic value type.
        /// </summary>
        /// <typeparam name="TOutput">The nullable type used to convert.</typeparam>
        /// <param name="input">The string to convert into the given generic type.</param>
        /// <returns>return null if the argument contains only empty chars or It is null.</returns>
        public static TOutput? ToPropertyFormat<TOutput>(string input)
            where TOutput : struct
        {
            input = NormalizeValue(input);

            if (input != null)
            {
                Type output = typeof(TOutput);
                if (IsDateTime(output))
                {
                    DateTime res;
                    if (DateTime.TryParseExact(input, DateFormats, DateTimeFormatInfo.CurrentInfo, DateTimeStyles.AssumeLocal, out res))
                        return (TOutput)Convert.ChangeType(res, output);
                }
                else
                {
                    return (TOutput)Convert.ChangeType(input, output);
                }
            }
            return (TOutput?)null;
        }

        /// <summary>
        /// Converts into string format
        /// </summary>
        /// <param name="obj">The object to convert into string format.</param>
        /// <returns>return null if the argument contains only empty chars or It is null.</returns>
        public static string ToExcelFormat(object obj)
        {
            return obj != null ? obj.ToString() : null;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public static string ToExcelFormat(DateTime? data)
        {
            return data.HasValue ? data.Value.ToShortDateString() : null;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="data"></param>
        /// <param name="format"></param>
        /// <param name="provider"></param>
        /// <returns></returns>
        public static string ToExcelFormat(DateTime? data, string format, IFormatProvider provider)
        {
            if (data.HasValue)
            {
                if (string.IsNullOrEmpty(format) && provider == null)
                    throw new ArgumentException("format provider or format argument must be not null.");

                bool isEmpty = string.IsNullOrEmpty(format);
                bool providerNull = provider == null;

                if (!isEmpty && provider != null)
                {
                    return data.Value.ToString(format, provider);
                }
                else if (!isEmpty)
                {
                    return data.Value.ToString(format);
                }
                else
                {
                    return data.Value.ToString(provider);
                }
            }
            return null;
        }
    }
}

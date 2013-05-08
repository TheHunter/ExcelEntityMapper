using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelEntityMapper.Exceptions
{
    /// <summary>
    /// 
    /// </summary>
    public class NotAvailableWorkSheetException
        :WorkBookException
    {
        private readonly string sheetName;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        public NotAvailableWorkSheetException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        /// <param name="sheetName"></param>
        public NotAvailableWorkSheetException(string message, string sheetName)
            : base(message)
        {
            this.sheetName = sheetName;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        public NotAvailableWorkSheetException(string message, Exception innerException)
            :base(message, innerException)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        /// <param name="sheetName"></param>
        /// <param name="innerException"></param>
        public NotAvailableWorkSheetException(string message, string sheetName, Exception innerException)
            : base(message, innerException)
        {
            this.sheetName = sheetName;
        }

        /// <summary>
        /// 
        /// </summary>
        public string SheetName
        {
            get { return this.sheetName; }
        }
    }
}

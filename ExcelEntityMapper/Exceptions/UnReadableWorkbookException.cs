using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelEntityMapper.Exceptions
{
    /// <summary>
    /// 
    /// </summary>
    public class UnReadableWorkbookException
        :WorkBookException
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        public UnReadableWorkbookException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        public UnReadableWorkbookException(string message, Exception innerException)
            :base(message, innerException)
        {
        }
    }
}

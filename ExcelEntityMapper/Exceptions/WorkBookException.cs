using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelEntityMapper.Exceptions
{
    /// <summary>
    /// 
    /// </summary>
    public class WorkBookException
        : Exception
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        public WorkBookException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        public WorkBookException(string message, Exception innerException)
            :base(message, innerException)
        {
        }
    }
}

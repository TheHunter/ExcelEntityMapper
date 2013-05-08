using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelEntityMapper.Exceptions
{
    /// <summary>
    /// 
    /// </summary>
    public class UnWriteableSheetException
        : WorkBookException
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        public UnWriteableSheetException(string message)
            : base(message)
        {
            
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        public UnWriteableSheetException(string message, Exception innerException)
            : base(message)
        {

        }
    }
}

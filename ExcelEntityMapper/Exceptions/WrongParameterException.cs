using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelEntityMapper.Exceptions
{
    /// <summary>
    /// 
    /// </summary>
    public class WrongParameterException
        : WorkBookException
    {
        private readonly string parameterName;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        /// <param name="parameterName"></param>
        public WrongParameterException(string message, string parameterName)
            : base(message)
        {
            this.parameterName = parameterName;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        /// <param name="parameterName"></param>
        /// <param name="innerException"></param>
        public WrongParameterException(string message, string parameterName, Exception innerException)
            :base(message, innerException)
        {
            this.parameterName = parameterName;
        }

        /// <summary>
        /// 
        /// </summary>
        public string ParameterName
        {
            get { return this.parameterName; }
        }
    }
}

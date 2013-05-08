using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelEntityMapper.Exceptions
{
    /// <summary>
    /// 
    /// </summary>
    public class SheetParameterException
        : WorkBookException
    {
        private string parameter;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        /// <param name="parameter"></param>
        public SheetParameterException(string message, string parameter)
            :base(message)
        {
            this.Parameter = parameter;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        /// <param name="parameter"></param>
        /// <param name="innerException"></param>
        public SheetParameterException(string message, string parameter, Exception innerException)
            :base(message, innerException)
        {
            this.Parameter = parameter;
        }

        /// <summary>
        /// 
        /// </summary>
        public string Parameter
        {
            private set
            {
                if (string.IsNullOrEmpty(value))
                    throw new WorkBookException("Parameter cannot be null.");
                this.parameter = value;
            }
            get { return this.parameter; }
        }
    }
}

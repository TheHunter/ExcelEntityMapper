using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelEntityMapper
{
    /// <summary>
    /// 
    /// </summary>
    /// <typeparam name="TSource"></typeparam>
    public interface IXLPropertyWriter<TSource>
        : IXLProperty
        where TSource : class
    {
        /// <summary>
        /// 
        /// </summary>
        Action<TSource, string> ToPropertyFormat { get; }
    }
}

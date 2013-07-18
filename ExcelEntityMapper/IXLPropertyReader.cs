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
    public interface IXLPropertyReader<TSource>
        : IXLProperty
        where TSource : class
    {
        /// <summary>
        /// 
        /// </summary>
        Func<TSource, string> ToExcelFormat { get; }
    }
}

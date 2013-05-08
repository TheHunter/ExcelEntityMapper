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
    public interface IXLSheetMapper<TSource>
        : IXLSheet
        where TSource : class, new()
    {
        /// <summary>
        /// 
        /// </summary>
        IEnumerable<IXLPropertyMapper<TSource>> PropertyMappers { get; }
    }
}

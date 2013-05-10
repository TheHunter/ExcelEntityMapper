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
        /// Executes the indicated action before mapping column values into properties instances.
        /// </summary>
        Action<TSource> BeforeMapping { get; }

        /// <summary>
        /// 
        /// </summary>
        IEnumerable<IXLPropertyMapper<TSource>> PropertyMappers { get; }
    }
}

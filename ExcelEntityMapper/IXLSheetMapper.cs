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
        /// Executes the indicated action after mapping column values into current computed instance.
        /// </summary>
        Action<TSource> AfterMapping { get; }

        /// <summary>
        /// A collection which contains custom instances for mapping properties with sheet columns.
        /// </summary>
        IEnumerable<IXLPropertyMapper<TSource>> PropertyMappers { get; }
    }
}

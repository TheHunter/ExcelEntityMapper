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
        where TSource : class//, new()
    {
        /// <summary>
        /// Executes the indicated action before setting instance values properties from worksheet assigned.
        /// </summary>
        Action<TSource> BeforeReading { get; set; }

        /// <summary>
        /// Executes the indicated action after setting instance values properties from worksheet assigned.
        /// </summary>
        Action<TSource> AfterReading { get; set; }
        
        /// <summary>
        /// Executes the indicated action before writing instance values properties into worksheet assigned.
        /// </summary>
        Action<TSource> BeforeWriting { get; set; }

        /// <summary>
        /// Executes the indicated action after writing instance values properties into worksheet assigned.
        /// </summary>
        Action<TSource> AfterWriting { get; set; }

        /// <summary>
        /// A collection which contains custom instances for mapping properties with sheet columns.
        /// </summary>
        IEnumerable<IXLPropertyMapper<TSource>> PropertyMappers { get; }
    }
}

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
        where TSource : class
    {
        /// <summary>
        /// A collection which contains custom instances for mapping properties with sheet columns.
        /// </summary>
        IEnumerable<IXLPropertyMapper<TSource>> PropertyMappers { get; }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="workbook"></param>
        void InjectWorkBook(IXLWorkBook workbook);
    }
}

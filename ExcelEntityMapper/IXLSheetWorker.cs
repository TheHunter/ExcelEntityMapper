using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelEntityMapper.Impl;

namespace ExcelEntityMapper
{
    /// <summary>
    /// Defines a methods for reading /writting worksheets, and a customizing property mappers.
    /// </summary>
    /// <typeparam name="TSource"></typeparam>
    public interface IXLSheetWorker<TSource>
        : IXLSheet ,IXLSheetReader<TSource>, IXLSheetWriter<TSource>
        where TSource : class
    {
        /// <summary>
        /// A set of customizing mappers for this WorksheetWorker.
        /// </summary>
        IEnumerable<IXLPropertyMapper<TSource>> PropertyMappers { get; }
    }
}

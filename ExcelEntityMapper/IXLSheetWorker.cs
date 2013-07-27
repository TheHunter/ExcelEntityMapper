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
        : IXLSheetReader<TSource>, IXLSheetWriter<TSource>
        where TSource : class, new()
    {
    }
}

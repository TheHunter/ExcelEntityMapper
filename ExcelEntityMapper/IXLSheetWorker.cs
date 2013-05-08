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
        : IXLSheetMapper<TSource>, IXLSheetReader<TSource>, IXLSheetWriter<TSource>
        where TSource : class, new()
    {

        /// <summary>
        /// 
        /// </summary>
        /// <param name="workbook"></param>
        void InjectWorkBook(IXLWorkBook workbook);
    }
}

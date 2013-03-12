using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelEntityMapper
{
    /// <summary>
    /// Defines a methods for reading /writting, filtering worksheets rows, and a customizing property mappers.
    /// </summary>
    /// <typeparam name="TSource">The source type which be used for transforming worksheet rows.</typeparam>
    public interface IXLSheetFiltered<TSource>
        : IXLSheetWorker<TSource>
        where TSource : class, new()
    {
        /// <summary>
        /// Filter the current worksheet rows by the given lambda expression.
        /// </summary>
        /// <param name="workbook">A workbook which has loaded the excel file.</param>
        /// <param name="buffer">A buffer which will be contain all intances read from worksheet.</param>
        /// <param name="function">A lambda expression for filtering instances read from the current worksheet.</param>
        /// <returns>Returns the row number read from the current worksheet.</returns>
        int ReadFilteredObjects(IXLWorkBook workbook, IDictionary<int, TSource> buffer, Func<TSource, bool> function);

        /// <summary>
        /// Filter the current worksheet rows by the given lambda expression.
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="sheetName">The sheetname to read from the current workbook.</param>
        /// <param name="buffer">A buffer which will be contain all intances read from worksheet.</param>
        /// <param name="function">A lambda expression for filtering instances read from the current worksheet.</param>
        /// <returns>Returns the row number read from the current worksheet.</returns>
        int ReadFilteredObjects(IXLWorkBook workbook, string sheetName, IDictionary<int, TSource> buffer, Func<TSource, bool> function);
    }
}

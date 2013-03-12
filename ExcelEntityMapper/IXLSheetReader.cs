using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelEntityMapper
{
    /// <summary>
    /// Defines a set of methods for reading custom objects into standard worksheets.
    /// </summary>
    /// <typeparam name="TSource">The source type which be used for transforming worksheet rows.</typeparam>
    public interface IXLSheetReader<TSource>
        where TSource : class
    {
        /// <summary>
        /// Reads all worksheet rows and transforms into objects.
        /// </summary>
        /// <param name="workbook">A workbook which has loaded the excel file.</param>
        /// <param name="buffer">A buffer which will be contain all intances read from worksheet.</param>
        /// <returns>Returns the row number read from the current worksheet.</returns>
        int ReadObjects(IXLWorkBook workbook, IDictionary<int, TSource> buffer);

        /// <summary>
        /// Reads all worksheet rows and transforms into objects.
        /// </summary>
        /// <param name="workbook">A workbook which has loaded the excel file.</param>
        /// <param name="sheetName">The sheetname to read from the current workbook.</param>
        /// <param name="buffer">A buffer which will be contain all intances read from worksheet.</param>
        /// <returns>Returns the row number read from the current worksheet.</returns>
        int ReadObjects(IXLWorkBook workbook, string sheetName, IDictionary<int, TSource> buffer);
    }
}

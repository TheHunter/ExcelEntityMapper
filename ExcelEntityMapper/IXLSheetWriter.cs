using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelEntityMapper
{
    /// <summary>
    /// Defines a set of methods for writing custom objects into standard worksheets.
    /// </summary>
    /// <typeparam name="TSource">The source type which be used for transforming worksheet rows.</typeparam>
    public interface IXLSheetWriter<TSource>
        where TSource : class
    {
        /// <summary>
        /// Writes all objects into current worksheet rows.
        /// </summary>
        /// <param name="instances">Objects which be written into current worksheet rows.</param>
        /// <returns>A number of objects written into the current worksheet.</returns>
        int WriteObjects(IEnumerable<TSource> instances);

        /// <summary>
        /// Writes all objects into current worksheet rows.
        /// </summary>
        /// <param name="sheetName">The sheetname to find for writing the current instances into workbook.</param>
        /// <param name="instances">Objects which be written into current worksheet rows.</param>
        /// <returns>A number of objects written into the current worksheet.</returns>
        int WriteObjects(string sheetName, IEnumerable<TSource> instances);
    }
}

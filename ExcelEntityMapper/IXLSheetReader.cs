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
        : IXLSheetMapper<TSource>
        where TSource : class, new()
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
        /// Tries to find the first readable row of the given sheet name.
        /// </summary>
        /// <param name="sheetName">the sheet name on looking for the first readable row.</param>
        /// <returns>returns the first readable row index of the internal worksheet, if no rows has found it returns -1.</returns>.
        int GetIndexFirstRow(string sheetName);

        /// <summary>
        /// Tries to read a row by the given index.
        /// </summary>
        /// <param name="rowIndex">the row index to read.</param>
        /// <returns>returns the object indicated in the given row index, if the given index doesn't corrispond to readable row, It returns null.</returns>
        TSource ReadObject(int rowIndex);

        /// <summary>
        /// Tries to read a row by the given index.
        /// </summary>
        /// <param name="sheetName">the sheet name for reading the row with the given index.</param>
        /// <param name="rowIndex">the row index to read.</param>
        /// <returns></returns>
        TSource ReadObject(string sheetName, int rowIndex);

        /// <summary>
        /// Reads all worksheet rows and transforms into objects.
        /// </summary>
        /// <param name="buffer">A buffer which will be contain all intances read from worksheet.</param>
        /// <returns>Returns the row number read from the current worksheet.</returns>
        int ReadObjects(IDictionary<int, TSource> buffer);

        /// <summary>
        /// Reads all worksheet rows and transforms into objects.
        /// </summary>
        /// <param name="sheetName">The sheetname to read from the current workbook.</param>
        /// <param name="buffer">A buffer which will be contain all intances read from worksheet.</param>
        /// <returns>Returns the row number read from the current worksheet.</returns>
        int ReadObjects(string sheetName, IDictionary<int, TSource> buffer);
    }
}

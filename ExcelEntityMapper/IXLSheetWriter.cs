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
        : IXLSheetMapper<TSource>
        where TSource : class
    {
        /// <summary>
        /// Executes the indicated action before writing instance values properties into worksheet assigned.
        /// </summary>
        Action<TSource> BeforeWriting { get; set; }

        /// <summary>
        /// Executes the indicated action after writing instance values properties into worksheet assigned.
        /// </summary>
        Action<TSource> AfterWriting { get; set; }

        /// <summary>
        /// Gets the last row index of the last row used.
        /// </summary>
        /// <param name="sheetName">The sheet name for looking for the last used row index.</param>
        /// <returns>returns the last used row index for the given sheet name, and It returns 0 if no rows are used.</returns>
        int GetIndexLastRow(string sheetName);

        /// <summary>
        /// Writes the given object into the next last row used present into the current worksheet.
        /// </summary>
        /// <param name="instance">the object to write on the current worksheet.</param>
        /// <returns>returns the next index of row which contains the given instance, and returns -1 if the instance wasn't be saved.</returns>
        int WriteObject(TSource instance);

        /// <summary>
        /// Writes the given object into the next last row used present into the current worksheet.
        /// </summary>
        /// <param name="sheetName">The sheet name on writing the given object.</param>
        /// <param name="instance">The object which will be saved into the current worksheet.</param>
        /// <returns>returns the next index of row which contains the given instance, and returns -1 if the instance wasn't be saved.</returns>
        int WriteObject(string sheetName, TSource instance);

        /// <summary>
        /// Writes the given object into the next last row used present into the current worksheet.
        /// </summary>
        /// <param name="rowIndex">The row index to use for saving the given instance.</param>
        /// <param name="instance">The object which will be saved into the current worksheet.</param>
        /// <returns>returns the next index of row which contains the given instance, and returns -1 if the instance wasn't be saved.</returns>
        int WriteObject(int rowIndex, TSource instance);

        /// <summary>
        /// Writes the given object into the next last row used present into the current worksheet.
        /// </summary>
        /// <param name="sheetName">The sheet name on writing the given object.</param>
        /// <param name="rowIndex">The row index to use for saving the given instance.</param>
        /// <param name="instance">The object which will be saved into the current worksheet.</param>
        /// <returns>returns the next index of row which contains the given instance, and returns -1 if the instance wasn't be saved.</returns>
        int WriteObject(string sheetName, int rowIndex, TSource instance);

        /// <summary>
        /// Writes all objects into current worksheet.
        /// </summary>
        /// <param name="instances">Objects which be written into current worksheet rows.</param>
        /// <returns>returns the next index of row which contains the given instance, and returns -1 if the instance wasn't be saved.</returns>
        int WriteObjects(IEnumerable<TSource> instances);

        /// <summary>
        /// Writes all objects into current worksheet rows.
        /// </summary>
        /// <param name="sheetName">The sheetname to find for writing the current instances into workbook.</param>
        /// <param name="instances">Objects which be written into current worksheet rows.</param>
        /// <returns>returns the next index of row which contains the given instance, and returns -1 if the instance wasn't be saved.</returns>
        int WriteObjects(string sheetName, IEnumerable<TSource> instances);

        /// <summary>
        /// Writes the row header on the spedified index.
        /// </summary>
        /// <param name="sheetName">the sheet name on writing the header.</param>
        /// <param name="rowIndex">the index on write the header.</param>
        void WriteHeader(string sheetName, int rowIndex);
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace ExcelEntityMapper
{
    /// <summary>
    /// 
    /// </summary>
    public interface IXLWorkBook
    {
        /// <summary>
        /// Checks if there's a worksheet with the given name.
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        bool ExistsWorkSheet(string sheetName);

        /// <summary>
        /// Returns all sheet names present into the calling workbook.
        /// </summary>
        /// <returns></returns>
        IEnumerable<string> GetSheetNames();

        /// <summary>
        /// Tries to add a new worksheet into the calling workbook.
        /// </summary>
        /// <param name="sheetName"></param>
        bool AddSheet(string sheetName);

        /// <summary>
        /// Tries to remove a worksheet with the given name.
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        bool RemoveSheet(string sheetName);

        /// <summary>
        /// 
        /// </summary>
        ExcelFormat Format { get; }

        /// <summary>
        /// Saves the currentstate of the calling workbook.
        /// </summary>
        /// <returns></returns>
        Stream Save();

        /// <summary>
        /// Creates a new workbook for the calling workbook, injecting the internal workbook into mapper workbook.
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="propertyMappers"></param>
        /// <returns></returns>
        IXLSheetWorker<TSource> MakeSheetMapper<TSource>(IEnumerable<IXLPropertyMapper<TSource>> propertyMappers)
            where TSource: class, new();
    }
}

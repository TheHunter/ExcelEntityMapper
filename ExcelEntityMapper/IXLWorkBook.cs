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
        /// 
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        bool ExistsWorkSheet(string sheetName);

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        IEnumerable<string> GetSheetNames();

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheetName"></param>
        bool AddSheet(string sheetName);

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        bool RemoveSheet(string sheetName);

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        Stream Save();
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelEntityMapper.Exceptions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace ExcelEntityMapper.Impl
{
    /// <summary>
    /// 
    /// </summary>
    internal static class MapperImplementor
    {
        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="wbProvider"></param>
        /// <param name="sheetName"></param>
        /// <param name="buffer"></param>
        /// <returns></returns>
        internal static int ReadObjects<TSource>(this IXWorkBookProvider<TSource> wbProvider, string sheetName, IDictionary<int, TSource> buffer)
            where TSource : class
        {
            HSSFWorkbook workBook = wbProvider.WorkBook;

            if (workBook == null)
                throw new UnReadableSheetException("The current workbook to use cannot be null.");

            ISheet workSheet = wbProvider.GetWorkSheet(sheetName);


            return 0;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="wbProvider"></param>
        /// <param name="sheetname"></param>
        /// <returns></returns>
        private static ISheet GetWorkSheet<TSource>(this IXWorkBookProvider<TSource> wbProvider, string sheetname)
            where TSource : class
        {
            HSSFWorkbook workBook = wbProvider.WorkBook;
            ISheet sheet = workBook.GetSheet(sheetname);
            if (sheet == null)
                throw new NotAvailableWorkSheetException("Impossible to find a worksheet with the specified name.", sheetname);

            return sheet;
        }
    }
}

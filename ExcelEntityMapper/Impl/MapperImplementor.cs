using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using ExcelEntityMapper.Exceptions;
using ExcelEntityMapper.Impl.BIFF;
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
        /// <param name="wbReader"></param>
        /// <param name="sheetName"></param>
        /// <param name="buffer"></param>
        /// <returns></returns>
        internal static int ReadObjects<TSource>(this IXWorkBookReader<TSource> wbReader, string sheetName, IDictionary<int, TSource> buffer)
            where TSource : class, new ()
        {
            HSSFWorkbook workBook = wbReader.WorkBook;

            if (workBook == null)
                throw new UnReadableSheetException("The current workbook to use cannot be null.");

            ISheet workSheet = wbReader.GetWorkSheet(sheetName);

            IRow row = wbReader.GetFirstRow(workSheet);
            int counter = 0;

            if (row == null)
                return counter;

            if (wbReader.HasHeader)
                row = workSheet.GetRow(row.RowNum + wbReader.HeaderRows);

            while (wbReader.IsReadableRow(row))
            {
                counter++;
                try
                {
                    TSource current = wbReader.ReadInstance(row);
                    if (current != null)
                        buffer.Add(row.RowNum, current);
                    row = workSheet.GetRow(row.RowNum + 1);
                }
                catch
                {
                    // any kind of exception cannot be blocking, so It reads the next row..
                }
            }
            return counter;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="wbReader"></param>
        /// <param name="workSheet"></param>
        /// <returns></returns>
        private static IRow GetFirstRow<TSource>(this IXWorkBookReader<TSource> wbReader, ISheet workSheet)
            where TSource : class, new()
        {
            int lastrow = workSheet.LastRowNum;
            IRow first = null;

            for (int index = 0; index <= lastrow; index++)
            {
                IRow current = workSheet.GetRow(index);
                if (wbReader.IsReadableRow(current))
                {
                    first = current;
                    break;
                }
            }
            return first;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="wbReader"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        private static bool IsReadableRow<TSource>(this IXWorkBookReader<TSource> wbReader, IRow row)
            where TSource : class, new()
        {
            if (row == null)
                return false;

            var col = wbReader.PropertyMappers.Where(n => n.CustomType == MapperType.Key);

            return col.Any() && col.All
                (
                    mapper =>
                    {
                        ICell cell = row.GetCell(mapper.ColumnIndex - wbReader.Offset);
                        if (cell == null || string.IsNullOrEmpty(cell.StringCellValue))
                            return false;
                        return true;
                    }
                );
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="wbReader"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        private static TSource ReadInstance<TSource>(this IXWorkBookReader<TSource> wbReader, IRow row)
            where TSource : class, new()
        {
            TSource instance = null;
            if (row != null)
            {
                instance = SourceHelper.CreateInstance<TSource>();

                if (wbReader.BeforeReading != null)
                    wbReader.BeforeReading.Invoke(instance);

                wbReader.PropertyMappers.All
                    (
                        parameter =>
                        {
                            try
                            {
                                ICell cell = row.GetCell(parameter.ColumnIndex - wbReader.Offset);
                                if (cell != null)
                                    parameter.ToPropertyFormat(instance, cell.StringCellValue);
                            }
                            catch
                            {
                                // nothing to do..
                            }
                            return true;
                        }
                    );

                if (wbReader.AfterReading != null)
                    wbReader.AfterReading(instance);
            }
            return instance;
            
        }


        public static int WriteObjects<TSource>(this IXWorkBookWriter<TSource> wbWriter , string sheetName, IEnumerable<TSource> instances)
            where TSource : class
        {
            HSSFWorkbook workBook = wbWriter.WorkBook;

            if (workBook == null)
                throw new UnWriteableSheetException("The current workbook to use cannot be null.");

            ISheet workSheet = wbWriter.GetWorkSheet(sheetName);

            return 0;
        }


        public static IRow GetLastRow(this IXWorkBookProvider wbProvider, ISheet workSheet)
        {
            //IRow last = wbProvider.GetFirstRow(workSheet);

            //if (last != null)
            //{
            //    int firstRow = last.RowNum;
            //    int lastrow = workSheet.LastRowNum;

            //    for (int index = firstRow; index <= lastrow; index++)
            //    {
            //        IRow tmp = workSheet.GetRow(index);
            //        if (!IsReadableRow(tmp))
            //            break;
            //        last = tmp;
            //    }
            //}
            //return last;
            return null;
        }

        #region Common methods
        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="wbReader"></param>
        /// <param name="sheetname"></param>
        /// <returns></returns>
        private static ISheet GetWorkSheet(this IXWorkBookProvider wbProvider, string sheetname)
        {
            HSSFWorkbook workBook = wbProvider.WorkBook;
            ISheet sheet = workBook.GetSheet(sheetname);
            if (sheet == null)
                throw new NotAvailableWorkSheetException("Impossible to find a worksheet with the specified name.", sheetname);

            return sheet;
        }
        #endregion
    }
}

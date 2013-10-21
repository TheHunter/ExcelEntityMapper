using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using ExcelEntityMapper.Exceptions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace ExcelEntityMapper.Impl.BIFF
{
    internal static class BiffMapperImplementor
    {
        private static readonly MethodInfo XlsNumericSetter;
        private static readonly MethodInfo XlsStringSetter;

        static BiffMapperImplementor()
        {
            XlsNumericSetter = typeof(ICell).GetMethod("SetCellValue", new[] { typeof(double) });
            XlsStringSetter = typeof(ICell).GetMethod("SetCellValue", new[] { typeof(string) });
        }

        #region BIFF sheet implementation

        #region Reader methods

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="wbReader"></param>
        /// <param name="sheetName"></param>
        /// <param name="buffer"></param>
        /// <returns>returns the number of rows read.</returns>
        internal static int ReadObjects<TSource>(this IXWorkBookReader<TSource> wbReader, string sheetName,
                                                 IDictionary<int, TSource> buffer)
            where TSource : class, new()
        {
            ISheet workSheet = wbReader.GetWorkSheet(sheetName);
            IRow row = wbReader.GetFirstRow(workSheet);

            if (row == null)
                return 0;

            if (wbReader.HasHeader)
                row = row.RowBelow(wbReader.HeaderRows);
            
            int counter = 0;
            while (wbReader.IsReadableRow(row))
            {
                counter++;
                try
                {
                    TSource current = wbReader.ReadInstance(row);
                    if (current != null)
                        buffer.Add(row.RowNumber() + wbReader.Offset, current);
                    row = row.RowBelow();
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
        /// <param name="sheetName"></param>
        /// <param name="rowIndex"></param>
        /// <returns></returns>
        internal static TSource ReadObject<TSource>(this IXWorkBookReader<TSource> wbReader, string sheetName,
                                                    int rowIndex)
            where TSource : class, new()
        {
            if ((rowIndex -= wbReader.Offset) < 0)
                throw new WrongParameterException(string.Format("The row index must be greater than zero, value: {0}", rowIndex), "indexRow");

            IRow row = wbReader.GetWorkSheet(sheetName)
                                .Row(rowIndex);

            return wbReader.IsReadableRow(row) ? wbReader.ReadInstance(row) : null;
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
                                ICell cell = row.Cell(parameter.ColumnIndex - wbReader.Offset);
                                parameter.ToPropertyFormat(instance, cell.GetString());
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

        #endregion

        #region Writer methods

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="wbWriter"></param>
        /// <param name="sheetName"></param>
        /// <param name="instance"></param>
        /// <returns></returns>
        internal static int WriteObject<TSource>(this IXWorkBookWriter<TSource> wbWriter, string sheetName,
                                                 TSource instance)
            where TSource : class
        {
            int rowIndex = wbWriter.GetIndexLastRow<TSource>(sheetName) + 1;
            return wbWriter.WriteObject<TSource>(sheetName, rowIndex, instance);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="wbWriter"></param>
        /// <param name="sheetName"></param>
        /// <param name="rowIndex"></param>
        /// <param name="instance"></param>
        /// <returns></returns>
        internal static int WriteObject<TSource>(this IXWorkBookWriter<TSource> wbWriter, string sheetName,
                                                 int rowIndex, TSource instance)
            where TSource : class
        {
            if ((rowIndex -= wbWriter.Offset) < 0)
                throw new WrongParameterException("The rowIndex parameters must be greater than zero in order to write data source into worksheet.", "rowIndex");

            IRow row = wbWriter.GetWorkSheet(sheetName)
                               .Row(rowIndex);

            if (wbWriter.WriteInstance(row, instance))
                return ++rowIndex;

            return -1;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="wbWriter"></param>
        /// <param name="sheetName"></param>
        /// <param name="instances"></param>
        /// <returns></returns>
        internal static int WriteObjects<TSource>(this IXWorkBookWriter<TSource> wbWriter, string sheetName,
                                                  IEnumerable<TSource> instances)
            where TSource : class
        {
            ISheet workSheet = wbWriter.GetWorkSheet(sheetName);

            int counter = 0;
            if (instances != null && instances.Any())
            {
                int rowIndex = -1;
                IRow lastRow = wbWriter.GetLastRow(workSheet);

                if (lastRow == null)
                {
                    if (wbWriter.HasHeader)
                    {
                        wbWriter.WriteHeader(++rowIndex, workSheet);
                    }
                }
                else
                {
                    rowIndex = lastRow.RowNumber();
                }

                foreach (var current in instances)
                {
                    IRow row = workSheet.Row(++rowIndex);
                    if (wbWriter.WriteInstance(row, current))
                        counter++;
                    else
                        rowIndex--;
                }
            }
            return counter;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="wbWriter"></param>
        /// <param name="rowIndex"></param>
        /// <param name="sheetName"></param>
        internal static void WriteHeader<TSource>(this IXWorkBookWriter<TSource> wbWriter, string sheetName,
                                                  int rowIndex)
            where TSource : class
        {
            if ((rowIndex -= wbWriter.Offset) < 0)
                throw new WrongParameterException("The rowIndex parameters must be greater than zero in order to write data source into worksheet.", "rowIndex");

            ISheet sheet = wbWriter.GetWorkSheet(sheetName);
            IRow row = sheet.Row(rowIndex);

            wbWriter.PropertyMappers.All
                (
                    parameter =>
                    {
                        row.Cell(parameter.ColumnIndex - wbWriter.Offset).SetCellValue(parameter.ColumnHeader);
                        return true;
                    }
                );
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="wbWriter"></param>
        /// <param name="rowBase"></param>
        /// <param name="sheet"></param>
        private static void WriteHeader<TSource>(this IXWorkBookWriter<TSource> wbWriter, int rowBase,
                                                 ISheet sheet)
            where TSource : class
        {
            var row = sheet.Row(rowBase);
            wbWriter.PropertyMappers.All
                (
                    parameter =>
                    {
                        row.Cell(parameter.ColumnIndex - wbWriter.Offset).SetCellValue(parameter.ColumnHeader);
                        return true;
                    }
                );
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="wbWriter"></param>
        /// <param name="row"></param>
        /// <param name="instance"></param>
        /// <returns></returns>
        private static bool WriteInstance<TSource>(this IXWorkBookWriter<TSource> wbWriter, IRow row,
                                                   TSource instance)
            where TSource : class
        {
            bool ret = false;

            if (wbWriter.BeforeWriting != null)
                wbWriter.BeforeWriting(instance);

            if (instance != null && row != null)
            {
                wbWriter.PropertyMappers.All
                (
                    parameter =>
                    {
                        try
                        {
                            ICell cell = row.Cell(parameter.ColumnIndex - wbWriter.Offset);
                            object val = XLEntityHelper.NormalizeXlsCellValue(parameter.ToExcelFormat(instance));
                            cell.SetCellValue(val);
                        }
                        catch (Exception)
                        {
                            //nothing to do..
                        }
                        return true;
                    }
                );
                ret = true;
            }

            if (wbWriter.AfterWriting != null)
                wbWriter.AfterWriting(instance);

            return ret;
        }

        #endregion

        #region Common methods

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="wbProvider"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        internal static int GetIndexFirstRow<TSource>(this IXWorkBookProvider<TSource> wbProvider, string sheetName)
            where TSource : class
        {
            ISheet workSheet = wbProvider.GetWorkSheet(sheetName);
            IRow row = wbProvider.GetFirstRow(workSheet);

            if (row == null)
                return -1;

            if (!wbProvider.HasHeader)
                return row.RowNumber() + wbProvider.Offset;

            row = row.RowBelow(wbProvider.HeaderRows);
            return wbProvider.IsReadableRow(row) ? row.RowNumber() + wbProvider.Offset : -1;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="wbProvider"></param>
        /// <param name="workSheet"></param>
        /// <returns></returns>
        private static IRow GetFirstRow<TSource>(this IXWorkBookProvider<TSource> wbProvider, ISheet workSheet)
            where TSource : class
        {
            int lastrow = workSheet.LastRowNum;
            IRow first = null;

            for (int index = 0; index <= lastrow; index++)
            {
                IRow current = workSheet.GetRow(index);
                if (wbProvider.IsReadableRow(current))
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
        /// <param name="wbProvider"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        private static bool IsReadableRow<TSource>(this IXWorkBookProvider<TSource> wbProvider, IRow row)
            where TSource : class
        {
            if (row == null)
                return false;

            var col = wbProvider.PropertyMappers.Where(n => n.CustomType == MapperType.Key);

            return col.Any() && col.All
                (
                    mapper =>
                    {
                        ICell cell = row.GetCell(mapper.ColumnIndex - wbProvider.Offset);
                        if (cell == null)
                            return false;

                        return !cell.IsEmpty();
                    }
                );
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="wbProvider"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        internal static int GetIndexLastRow<TSource>(this IXWorkBookProvider<TSource> wbProvider, string sheetName)
            where TSource : class
        {
            ISheet workSheet = wbProvider.GetWorkSheet(sheetName);
            IRow lastRow = wbProvider.GetLastRow(workSheet);
            if (lastRow != null)
                return lastRow.RowNum + wbProvider.Offset;

            return wbProvider.HeaderRows;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="wbProvider"></param>
        /// <param name="workSheet"></param>
        /// <returns></returns>
        private static IRow GetLastRow<TSource>(this IXWorkBookProvider<TSource> wbProvider, ISheet workSheet)
            where TSource : class
        {
            IRow last = wbProvider.GetFirstRow(workSheet);

            if (last != null)
            {
                int firstRow = last.RowNum;
                int lastrow = workSheet.LastRowNum;

                for (int index = firstRow; index <= lastrow; index++)
                {
                    IRow tmp = workSheet.GetRow(index);
                    if (!wbProvider.IsReadableRow(tmp))
                        break;
                    last = tmp;
                }
            }
            return last;

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

            if (workBook == null)
                throw new UnReadableSheetException("The current XWorkBook to use cannot be null.");

            ISheet sheet = workBook.GetSheet(sheetname);
            if (sheet == null)
                throw new NotAvailableWorkSheetException("Impossible to find a worksheet with the specified name.", sheetname);

            return sheet;
        }

        #endregion

        #region Other

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="wbProvider"></param>
        /// <param name="workbook"></param>
        /// <exception cref="SheetParameterException"></exception>
        /// <exception cref="WorkBookException"></exception>
        internal static void InjectWorkBook<TSource>(this IXWorkBookProvider<TSource> wbProvider, IXLWorkBook workbook)
            where TSource : class
        {
            if (workbook == null)
                throw new WorkBookException("Impossible to associate a null WorkBook.");

            XWorkBook current = workbook as XWorkBook;
            if (current == null)
                throw new SheetParameterException("The WorkBook instance isn't valid format.", "WorkBook");

            wbProvider.WorkBook = current.WorkBook;
        }
        #endregion

        #endregion
    }
}

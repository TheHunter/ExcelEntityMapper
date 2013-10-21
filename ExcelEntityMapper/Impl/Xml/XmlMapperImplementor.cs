using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using ExcelEntityMapper.Exceptions;

namespace ExcelEntityMapper.Impl.Xml
{
    internal static class XmlMapperImplementor
    {
        #region XML sheet implementation

        #region Reader methods

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="wbReader"></param>
        /// <param name="sheetName"></param>
        /// <param name="buffer"></param>
        /// <returns></returns>
        internal static int ReadObjects<TSource>(this IXLWorkBookReader<TSource> wbReader, string sheetName,
                                                 IDictionary<int, TSource> buffer)
            where TSource : class, new()
        {
            IXLWorksheet workSheet = wbReader.GetWorkSheet(sheetName);
            IXLRow row = wbReader.GetFirstRow(workSheet);

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
                        buffer.Add(row.RowNumber(), current);
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
        internal static TSource ReadObject<TSource>(this IXLWorkBookReader<TSource> wbReader, string sheetName,
                                                    int rowIndex)
            where TSource : class, new()
        {
            if (rowIndex < 1)
                throw new WrongParameterException(string.Format("The row index must be greater than zero, value: {0}", rowIndex), "indexRow");

            IXLRow row = wbReader.GetWorkSheet(sheetName)
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
        private static TSource ReadInstance<TSource>(this IXLWorkBookReader<TSource> wbReader, IXLRow row)
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
                                IXLCell cell = row.Cell(parameter.ColumnIndex);
                                parameter.ToPropertyFormat(instance, cell.GetString().Trim());
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
        internal static int WriteObject<TSource>(this IXLWorkBookWriter<TSource> wbWriter, string sheetName,
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
        internal static int WriteObject<TSource>(this IXLWorkBookWriter<TSource> wbWriter, string sheetName,
                                                 int rowIndex, TSource instance)
            where TSource : class
        {
            if (rowIndex < 1)
                throw new WrongParameterException("The rowIndex parameters must be greater than zero in order to write data source into worksheet.", "rowIndex");

            IXLRow row = wbWriter.GetWorkSheet(sheetName)
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
        internal static int WriteObjects<TSource>(this IXLWorkBookWriter<TSource> wbWriter, string sheetName,
                                                  IEnumerable<TSource> instances)
            where TSource : class
        {
            IXLWorksheet workSheet = wbWriter.GetWorkSheet(sheetName);

            int counter = 0;
            if (instances != null && instances.Any())
            {
                int rowIndex = 0;
                IXLRow lastRow = wbWriter.GetLastRow(workSheet);

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
                    IXLRow row = workSheet.Row(++rowIndex);
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
        internal static void WriteHeader<TSource>(this IXLWorkBookWriter<TSource> wbWriter, string sheetName,
                                                  int rowIndex)
            where TSource : class
        {
            if ((rowIndex -= wbWriter.Offset) < 0)
                throw new WrongParameterException("The rowIndex parameters must be greater than zero in order to write data source into worksheet.", "rowIndex");

            IXLWorksheet sheet = wbWriter.GetWorkSheet(sheetName);
            IXLRow row = sheet.Row(rowIndex);

            wbWriter.PropertyMappers.All
                (
                    parameter =>
                    {
                        row.Cell(parameter.ColumnIndex).Value = parameter.ColumnHeader;
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
        private static void WriteHeader<TSource>(this IXLWorkBookWriter<TSource> wbWriter, int rowBase,
                                                 IXLWorksheet sheet)
            where TSource : class
        {
            var row = sheet.Row(rowBase);
            wbWriter.PropertyMappers.All
                (
                    parameter =>
                    {
                        row.Cell(parameter.ColumnIndex).Value = parameter.ColumnHeader;
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
        private static bool WriteInstance<TSource>(this IXLWorkBookWriter<TSource> wbWriter, IXLRow row,
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
                            IXLCell cell = row.Cell(parameter.ColumnIndex);
                            cell.Value = parameter.ToExcelFormat(instance);
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
        /// <param name="wbReader"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        internal static int GetIndexFirstRow<TSource>(this IXLWorkBookProvider<TSource> wbReader, string sheetName)
            where TSource : class
        {
            IXLWorksheet workSheet = wbReader.GetWorkSheet(sheetName);
            IXLRow row = wbReader.GetFirstRow(workSheet);

            if (row == null)
                return -1;

            if (!wbReader.HasHeader)
                return row.RowNumber();

            row = row.RowBelow(wbReader.HeaderRows);
            return wbReader.IsReadableRow(row) ? row.RowNumber() : -1;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="wbProvider"></param>
        /// <param name="workSheet"></param>
        /// <returns></returns>
        private static IXLRow GetFirstRow<TSource>(this IXLWorkBookProvider<TSource> wbProvider, IXLWorksheet workSheet)
            where TSource : class
        {
            IXLRow firstRow = null;
            IXLRow lastRow = workSheet.LastRowUsed();

            int firstIndex = 1;
            if (lastRow != null)
            {
                int lastIndex = lastRow.RowNumber();

                for (int index = firstIndex; index <= lastIndex; index++)
                {
                    lastRow = workSheet.Row(index);
                    if (wbProvider.IsReadableRow(lastRow))
                    {
                        firstRow = lastRow;
                        break;
                    }
                }
            }

            return firstRow;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="wbProvider"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        private static bool IsReadableRow<TSource>(this IXLWorkBookProvider<TSource> wbProvider, IXLRow row)
            where TSource : class
        {
            var col = wbProvider.PropertyMappers.Where(n => n.CustomType == MapperType.Key);
            return col.Any() && col.All(n => !row.Cell(n.ColumnIndex).IsEmpty());
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="wbProvider"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        internal static int GetIndexLastRow<TSource>(this IXLWorkBookProvider<TSource> wbProvider, string sheetName)
            where TSource : class
        {
            IXLWorksheet workSheet = wbProvider.GetWorkSheet(sheetName);
            IXLRow lastRow = wbProvider.GetLastRow(workSheet);

            if (lastRow != null)
                return lastRow.RowNumber();

            return wbProvider.HeaderRows;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="wbProvider"></param>
        /// <param name="workSheet"></param>
        /// <returns></returns>
        private static IXLRow GetLastRow<TSource>(this IXLWorkBookProvider<TSource> wbProvider, IXLWorksheet workSheet)
            where TSource : class
        {
            IXLRow firstRow = wbProvider.GetFirstRow(workSheet);
            IXLRow lastRow = workSheet.LastRowUsed();

            if (firstRow != null && lastRow != null)
            {
                int firstIndex = firstRow.RowNumber();
                int lastIndex = lastRow.RowNumber();

                for (int index = firstIndex; index <= lastIndex; index++)
                {
                    IXLRow currRow = workSheet.Row(index);
                    if (!wbProvider.IsReadableRow(currRow))
                        break;
                    lastRow = currRow;
                }
            }
            else
            {
                lastRow = null;
            }

            return lastRow;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="wbProvider"></param>
        /// <param name="sheetname"></param>
        /// <returns></returns>
        private static IXLWorksheet GetWorkSheet<TSource>(this IXLWorkBookProvider<TSource> wbProvider, string sheetname)
            where TSource : class
        {
            if (wbProvider.WorkBook == null)
                throw new UnReadableSheetException("The current XLWorkBook to use cannot be null.");

            try
            {
                return wbProvider.WorkBook.Worksheet(sheetname);
            }
            catch (Exception ex)
            {
                throw new NotAvailableWorkSheetException("Impossible to find a worksheet with the specified name.", sheetname, ex);
            }
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
        internal static void InjectWorkBook<TSource>(this IXLWorkBookProvider<TSource> wbProvider, IXLWorkBook workbook)
            where TSource : class
        {
            if (workbook == null)
                throw new WorkBookException("Impossible to associate a null WorkBook.");

            XLWorkBook current = workbook as XLWorkBook;
            if (current == null)
                throw new SheetParameterException("The WorkBook instance isn't valid format.", "WorkBook");

            wbProvider.WorkBook = current.WorkBook;
        }
        #endregion

        #endregion
    }
}

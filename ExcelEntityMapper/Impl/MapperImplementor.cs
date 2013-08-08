﻿using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using ClosedXML.Excel;
using ExcelEntityMapper.Exceptions;
using ExcelEntityMapper.Impl.BIFF;
using ExcelEntityMapper.Impl.Xml;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace ExcelEntityMapper.Impl
{
    /// <summary>
    /// Extension class for sheet mappers.
    /// </summary>
    public static class MapperImplementor
    {
        private static readonly MethodInfo XlsNumericSetter;
        private static readonly MethodInfo XlsStringSetter;

        static MapperImplementor()
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
        /// <returns></returns>
        internal static int ReadObjects<TSource>(this IXWorkBookReader<TSource> wbReader, string sheetName, IDictionary<int, TSource> buffer)
            where TSource : class, new()
        {
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
        /// <param name="sheetName"></param>
        /// <param name="indexRow"></param>
        /// <returns></returns>
        internal static TSource ReadObject<TSource>(this IXWorkBookReader<TSource> wbReader, string sheetName, int indexRow)
            where TSource : class, new()
        {
            indexRow -= wbReader.Offset;

            if (indexRow < 0)
                throw new WrongParameterException(string.Format("The index row must be greater than zero, value: {0}", indexRow), "indexRow");

            ISheet workSheet = wbReader.GetWorkSheet(sheetName);
            return wbReader.ReadInstance(workSheet.GetRow(indexRow));
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
                                {
                                    string cellValue = null;
                                    cellValue = cell.CellType == CellType.NUMERIC ? cell.NumericCellValue.ToString(CultureInfo.InvariantCulture) : cell.StringCellValue;
                                    parameter.ToPropertyFormat(instance, cellValue);
                                }
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
        /// <param name="instances"></param>
        /// <returns></returns>
        internal static int WriteObjects<TSource>(this IXWorkBookWriter<TSource> wbWriter, string sheetName, IEnumerable<TSource> instances)
            where TSource : class
        {
            //HSSFWorkbook workBook = wbWriter.WorkBook;
            //if (workBook == null)
            //    throw new UnWriteableSheetException("The current WorkBook to use cannot be null.");

            ISheet workSheet = wbWriter.GetWorkSheet(sheetName);

            int counter = 0;
            if (instances != null && instances.Any())
            {
                int rowIndex = -1;
                IRow lastRow = wbWriter.GetLastRow(workSheet);

                if (wbWriter.HasHeader)
                {
                    if (lastRow == null)
                    {
                        rowIndex = 0;
                        wbWriter.WriteHeader(rowIndex, workSheet);
                    }
                    else
                    {
                        rowIndex = lastRow.RowNum; //+ (this.HeaderRows - 1);
                    }
                }

                foreach (var current in instances)
                {
                    rowIndex++;
                    IRow row = workSheet.GetRow(rowIndex) ?? workSheet.CreateRow(rowIndex);

                    if (wbWriter.WriteInstance(row, current))
                    {
                        counter++;
                    }
                }
            }
            return counter;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="wbWriter"></param>
        /// <param name="rowBase"></param>
        /// <param name="sheet"></param>
        private static void WriteHeader<TSource>(this IXWorkBookWriter<TSource> wbWriter, int rowBase, ISheet sheet)
            where TSource : class
        {
            int index = rowBase;
            var row = sheet.GetRow(index);
            if (row == null)
            {
                row = sheet.CreateRow(index);
                wbWriter.PropertyMappers.All
                (
                    delegate(IXLPropertyMapper<TSource> parameter)
                    {
                        row.CreateCell(parameter.ColumnIndex - wbWriter.Offset)
                            .SetCellValue(parameter.ColumnHeader);
                        return true;
                    }
                );
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="wbWriter"></param>
        /// <param name="row"></param>
        /// <param name="instance"></param>
        /// <returns></returns>
        private static bool WriteInstance<TSource>(this IXWorkBookWriter<TSource> wbWriter, IRow row, TSource instance)
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
                            ICell cell = row.CreateCell(parameter.ColumnIndex - wbWriter.Offset);
                            object val = XLEntityHelper.NormalizeXlsCellValue(parameter.ToExcelFormat(instance));
                            object[] parameters = { val };

                            if (XLEntityHelper.IsNumericValue(val))
                                XlsNumericSetter.Invoke(cell, parameters);
                            else
                                XlsStringSetter.Invoke(cell, parameters);
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
        internal static int GetIndexFirstRow<TSource>(this IXWorkBookProvider<TSource> wbReader, string sheetName)
            where TSource : class
        {
            ISheet workSheet = wbReader.GetWorkSheet(sheetName);
            IRow row = wbReader.GetFirstRow(workSheet);

            if (row == null)
                return -1;

            int indexRow = row.RowNum;
            if (wbReader.HasHeader)
                indexRow += wbReader.HeaderRows;

            row = workSheet.GetRow(indexRow);
            return wbReader.IsReadableRow(row) ? row.RowNum + wbReader.Offset : -1;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="wbReader"></param>
        /// <param name="workSheet"></param>
        /// <returns></returns>
        private static IRow GetFirstRow<TSource>(this IXWorkBookProvider<TSource> wbReader, ISheet workSheet)
            where TSource : class
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
        private static bool IsReadableRow<TSource>(this IXWorkBookProvider<TSource> wbReader, IRow row)
            where TSource : class
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
        internal static int ReadObjects<TSource>(this IXLWorkBookReader<TSource> wbReader, string sheetName, IDictionary<int, TSource> buffer)
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
        /// <param name="indexRow"></param>
        /// <returns></returns>
        internal static TSource ReadObject<TSource>(this IXLWorkBookReader<TSource> wbReader, string sheetName, int indexRow)
            where TSource : class, new()
        {
            if (indexRow < 1)
                throw new WrongParameterException(string.Format("The index row must be greater than zero, value: {0}", indexRow), "indexRow");

            IXLWorksheet workSheet = wbReader.GetWorkSheet(sheetName);
            return wbReader.ReadInstance(workSheet.Row(indexRow));
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

                if (wbWriter.HasHeader)
                {
                    if (lastRow == null)
                    {
                        rowIndex = 1;
                        wbWriter.WriteHeader(rowIndex, workSheet);
                    }
                    else
                    {
                        rowIndex = lastRow.RowNumber();
                    }
                }

                foreach (var current in instances)
                {
                    rowIndex++;
                    IXLRow row = workSheet.Row(rowIndex);
                    if (wbWriter.WriteInstance(row, current))
                        counter++;
                }
            }
            return counter;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="wbWriter"></param>
        /// <param name="rowBase"></param>
        /// <param name="sheet"></param>
        private static void WriteHeader<TSource>(this IXLWorkBookWriter<TSource> wbWriter, int rowBase, IXLWorksheet sheet)
            where TSource : class
        {
            var row = sheet.Row(rowBase);

            wbWriter.PropertyMappers.All
                (
                    delegate(IXLPropertyMapper<TSource> parameter)
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
        private static bool WriteInstance<TSource>(this IXLWorkBookWriter<TSource> wbWriter, IXLRow row, TSource instance)
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
        internal static int GetIndexFirstRow<TSource>(this IXLWorkBookProvider<TSource> wbReader, string sheetName)
            where TSource : class
        {
            IXLWorksheet workSheet = wbReader.GetWorkSheet(sheetName);
            IXLRow row = wbReader.GetFirstRow(workSheet);

            if (row == null)
                return -1;

            if (wbReader.HasHeader)
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

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="sheetMapper"></param>
        /// <param name="sheetName"></param>
        /// <param name="buffer"></param>
        /// <param name="function"></param>
        /// <returns></returns>
        internal static int ReadFilteredObjects<TSource>(this IXLSheetFiltered<TSource> sheetMapper, string sheetName, IDictionary<int, TSource> buffer,
                                                Func<TSource, bool> function)
             where TSource : class, new()
        {
            if (buffer == null)
                throw new SheetParameterException("Buffer cannot be null.", "buffer");

            Dictionary<int, TSource> temp = new Dictionary<int, TSource>();
            int count = sheetMapper.ReadObjects(sheetName, temp);

            var filtered = temp.Where(n => function(n.Value));

            if (filtered.Any())
            {
                foreach (var current in filtered)
                {
                    buffer.Add(current);
                }
            }

            return count;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="sheetWorker"></param>
        public static IXLSheetWorker<TSource> AsXmlSheetWorker<TSource>(this IXLSheetWorker<TSource> sheetWorker)
            where TSource : class, new()
        {
            return new XLSheetMapper<TSource>(sheetWorker.HeaderRows, sheetWorker.PropertyMappers);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="sheetWorker"></param>
        /// <returns></returns>
        public static IXLSheetMapper<TSource> AsBiffSheetWorker<TSource>(this IXLSheetWorker<TSource> sheetWorker)
            where TSource : class, new()
        {
            return new XSheetMapper<TSource>(sheetWorker.HeaderRows, sheetWorker.PropertyMappers);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="sheetWorker"></param>
        /// <returns></returns>
        public static IXLSheetReader<TSource> AsXmlReader<TSource>(this IXLSheetWorker<TSource> sheetWorker)
            where TSource : class, new()
        {
            return new XLSheetReader<TSource>(sheetWorker.HeaderRows, sheetWorker.PropertyMappers);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="sheetWorker"></param>
        /// <returns></returns>
        public static IXLSheetReader<TSource> AsBiffReader<TSource>(this IXLSheetWorker<TSource> sheetWorker)
            where TSource : class, new()
        {
            return new XSheetReader<TSource>(sheetWorker.HeaderRows, sheetWorker.PropertyMappers);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="sheetWorker"></param>
        /// <returns></returns>
        public static IXLSheetWriter<TSource> AsXmlWriter<TSource>(this IXLSheetWorker<TSource> sheetWorker)
            where TSource : class, new()
        {
            return new XLSheetWriter<TSource>(sheetWorker.HeaderRows, sheetWorker.PropertyMappers);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="sheetWorker"></param>
        /// <returns></returns>
        public static IXLSheetWriter<TSource> AsBiffWriter<TSource>(this IXLSheetWorker<TSource> sheetWorker)
            where TSource : class, new()
        {
            return new XSheetWriter<TSource>(sheetWorker.HeaderRows, sheetWorker.PropertyMappers);
        }
    }
}

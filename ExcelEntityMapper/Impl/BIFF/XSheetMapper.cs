using System;
using System.Collections.Generic;
using System.Linq;
using ExcelEntityMapper.Exceptions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace ExcelEntityMapper.Impl.BIFF
{
    /// <summary>
    /// 
    /// </summary>
    /// <typeparam name="TSource"></typeparam>
    public class XSheetMapper<TSource>
        : SheetMapper<TSource>, IXLSheetWorker<TSource>
        where TSource : class, new()
    {
        private HSSFWorkbook workBook;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="indexkeyColumn"></param>
        /// <param name="propertyMappers"></param>
        public XSheetMapper(int indexkeyColumn, IEnumerable<IXLPropertyMapper<TSource>> propertyMappers)
            : this(indexkeyColumn, false, propertyMappers)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="indexkeyColumn"></param>
        /// <param name="hasHeader"></param>
        /// <param name="propertyMappers"></param>
        public XSheetMapper(int indexkeyColumn, bool hasHeader, IEnumerable<IXLPropertyMapper<TSource>> propertyMappers)
            : base(indexkeyColumn, hasHeader, true, propertyMappers)
        {
            this.LastColumn = this.PropertyMappers.Select(n => n.ColumnIndex).Max() - this.Offset;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="rowBase"></param>
        /// <param name="sheet"></param>
        protected void WriteHeader(int rowBase, ISheet sheet)
        {
            int index = rowBase;
            var row = sheet.GetRow(index);
            if (row == null)
            {
                row = sheet.CreateRow(index);
                this.PropertyMappers.All
                (
                    delegate(IXLPropertyMapper<TSource> parameter)
                    {
                        row.CreateCell(parameter.ColumnIndex - this.Offset)
                            .SetCellValue(parameter.ColumnHeader);
                        return true;
                    }
                );
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="workbook"></param>
        public void InjectWorkBook(IXLWorkBook workbook)
        {
            if (workbook == null)
                throw new WorkBookException("Impossible to associate a null workbook.");

            XWorkBook current = workbook as XWorkBook;
            if (current == null)
                throw new SheetParameterException("The workbook instance isn't valid format.", "workbook");

            this.workBook = current.Workbook;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="buffer"></param>
        /// <returns></returns>
        public int ReadObjects(IDictionary<int, TSource> buffer)
        {
            return this.ReadObjects(this.SheetName, buffer);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="buffer"></param>
        /// <returns></returns>
        public int ReadObjects(string sheetName, IDictionary<int, TSource> buffer)
        {
            if (this.workBook == null)
                throw new UnReadableSheetException("The current workbook to use cannot be null.");

            ISheet workSheet = this.GetWorkSheet(sheetName);

            IRow row = this.GetFirstRow(workSheet);
            int counter = 0;

            if (row == null)        //no rows used.
                return counter;

            if (this.HasHeader)
                row = workSheet.GetRow(row.RowNum + 1);

            ICell KeyCell;
            while (row != null
                && (KeyCell = row.GetCell(this.IndexKeyColumn - this.Offset)) != null
                && !string.IsNullOrEmpty(KeyCell.StringCellValue)
                )
            {
                counter++;
                TSource current = this.ReadInstance(row);
                if (current != null)
                {
                    buffer.Add(row.RowNum, current);
                }
                row = workSheet.GetRow(row.RowNum + 1);
            }
            return counter;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        private TSource ReadInstance(IRow row)
        {
            TSource instance = null;
            if (row != null)
            {
                instance = SourceHelper.CreateInstance<TSource>();

                if (this.BeforeMapping != null)
                    this.BeforeMapping.Invoke(instance);

                this.PropertyMappers.All
                    (
                        parameter =>
                        {
                            try
                            {
                                ICell cell = row.GetCell(parameter.ColumnIndex - this.Offset);
                                if (cell != null)
                                {
                                    parameter.ToPropertyFormat(instance, cell.StringCellValue);
                                }
                            }
                            catch
                            {
                                // nothing to do..
                            }
                            return true;
                        }
                    );
            }
            return instance;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="instances"></param>
        /// <returns></returns>
        public int WriteObjects(IEnumerable<TSource> instances)
        {
            return this.WriteObjects(this.SheetName, instances);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="instances"></param>
        /// <returns></returns>
        public int WriteObjects(string sheetName, IEnumerable<TSource> instances)
        {
            if (this.workBook == null)
                throw new UnWriteableSheetException("The current workbook to use cannot be null.");

            ISheet workSheet = this.GetWorkSheet(sheetName);

            int counter = 0;
            if (instances != null && instances.Any())
            {
                IRow row;
                int rowIndex = -1;
                IRow header = GetFirstRow(workSheet);

                if (this.HasHeader)
                {
                    if (header == null)
                    {
                        rowIndex = 0;
                        this.WriteHeader(rowIndex, workSheet);
                    }
                    else
                    {
                        rowIndex = header.RowNum;
                    }
                }

                foreach (var current in instances)
                {
                    rowIndex++;
                    row = workSheet.GetRow(rowIndex);

                    if (row == null)
                        row = workSheet.CreateRow(rowIndex);

                    if (this.WriteInstance(row, current))
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
        /// <param name="row"></param>
        /// <param name="instance"></param>
        /// <returns></returns>
        private bool WriteInstance(IRow row, TSource instance)
        {
            bool ret = false;
            if (instance != null && row != null)
            {
                ICell cell;
                this.PropertyMappers.All
                (
                    parameter =>
                    {
                        try
                        {
                            cell = row.CreateCell(parameter.ColumnIndex - this.Offset);
                            cell.SetCellValue(parameter.ToExcelFormat(instance));
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
            return ret;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheetname"></param>
        /// <returns></returns>
        private ISheet GetWorkSheet(string sheetname)
        {
            ISheet sheet = this.workBook.GetSheet(sheetname);
            if (sheet == null)
                throw new NotAvailableWorkSheetException("Impossible to find a worksheet with the specified name.", sheetname);
            
            return sheet;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="workSheet"></param>
        /// <returns></returns>
        private IRow GetFirstRow(ISheet workSheet)
        {
            int firstRow = workSheet.FirstRowNum;
            int lastrow = workSheet.LastRowNum;
            int indKey = this.IndexKeyColumn - this.Offset;
            
            IRow first = null;

            for (int index = firstRow; index <= lastrow; index++)
            {
                first = workSheet.GetRow(index);
                if (first != null && first.Cells.FirstOrDefault(n => n.ColumnIndex == indKey) != null)
                    break;
            }
            return first;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="workSheet"></param>
        /// <returns></returns>
        private IRow GetLastRow(ISheet workSheet)
        {
            IRow last = this.GetFirstRow(workSheet);
            IRow tmp = null;

            if (last != null)
            {
                int firstRow = last.RowNum + 1;
                int lastrow = workSheet.LastRowNum;
                int indKey = this.IndexKeyColumn - this.Offset;

                for (int index = firstRow; index <= lastrow; index++)
                {
                    tmp = workSheet.GetRow(index);
                    if (tmp == null || tmp.Cells.FirstOrDefault(n => n.ColumnIndex == indKey) == null)
                    {
                        break;
                    }
                    last = tmp;
                }
            }
            return last;
        }
    }
}

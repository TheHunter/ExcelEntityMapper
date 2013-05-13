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
        /// <param name="propertyMappers"></param>
        public XSheetMapper(IEnumerable<IXLPropertyMapper<TSource>> propertyMappers)
            : this(0, propertyMappers)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="headerRows"></param>
        /// <param name="propertyMappers"></param>
        public XSheetMapper(int headerRows, IEnumerable<IXLPropertyMapper<TSource>> propertyMappers)
            : base(headerRows, true, propertyMappers)
        {
            //this.LastColumn = this.PropertyMappers.Select(n => n.ColumnIndex).Max() - this.Offset;
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
        /// <param name="sheetName"></param>
        /// <param name="buffer"></param>
        /// <returns></returns>
        public override int ReadObjects(string sheetName, IDictionary<int, TSource> buffer)
        {
            if (this.workBook == null)
                throw new UnReadableSheetException("The current workbook to use cannot be null.");

            ISheet workSheet = this.GetWorkSheet(sheetName);

            IRow row = this.GetFirstRow(workSheet);
            int counter = 0;

            if (row == null)
                return counter;

            if (this.HasHeader)
                row = workSheet.GetRow(row.RowNum + this.HeaderRows);

            while (IsReadableRead(row))
            {
                counter++;
                try
                {
                    TSource current = this.ReadInstance(row);
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
        /// <param name="row"></param>
        /// <returns></returns>
        private TSource ReadInstance(IRow row)
        {
            TSource instance = null;
            if (row != null)
            {
                instance = SourceHelper.CreateInstance<TSource>();

                if (this.BeforeReading != null)
                    this.BeforeReading.Invoke(instance);

                this.PropertyMappers.All
                    (
                        parameter =>
                        {
                            try
                            {
                                ICell cell = row.GetCell(parameter.ColumnIndex - this.Offset);
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

                if (this.AfterReading != null)
                    this.AfterReading(instance);
            }
            return instance;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        private bool IsReadableRead(IRow row)
        {
            if (row == null)
                return false;

            var col = this.PropertyMappers.Where(n => n.CustomType == MapperType.Key);
            return col.Any() && col.All
                (
                    mapper =>
                        {
                            ICell cell = row.GetCell(mapper.ColumnIndex - this.Offset);
                            if (cell == null || string.IsNullOrEmpty(cell.StringCellValue))
                                return false;
                            return true;
                        }
                );
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="instances"></param>
        /// <returns></returns>
        public override int WriteObjects(string sheetName, IEnumerable<TSource> instances)
        {
            if (this.workBook == null)
                throw new UnWriteableSheetException("The current workbook to use cannot be null.");

            ISheet workSheet = this.GetWorkSheet(sheetName);

            int counter = 0;
            if (instances != null && instances.Any())
            {
                int rowIndex = -1;
                IRow lastRow = GetLastRow(workSheet); //

                if (this.HasHeader)
                {
                    if (lastRow == null)
                    {
                        rowIndex = 0;
                        this.WriteHeader(rowIndex, workSheet);
                    }
                    else
                    {
                        rowIndex = lastRow.RowNum; //+ (this.HeaderRows - 1);
                    }
                }

                foreach (var current in instances)
                {
                    rowIndex++;
                    IRow row = workSheet.GetRow(rowIndex);

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

            if (this.BeforeWriting != null)
                this.BeforeWriting(instance);

            if (instance != null && row != null)
            {
                this.PropertyMappers.All
                (
                    parameter =>
                    {
                        try
                        {
                            ICell cell = row.CreateCell(parameter.ColumnIndex - this.Offset);
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

            if (this.AfterWriting != null)
                this.AfterWriting(instance);

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
            
            IRow first = null;

            for (int index = firstRow; index <= lastrow; index++)
            {
                first = workSheet.GetRow(index);
                //if (first != null && first.Cells.FirstOrDefault(n => n.ColumnIndex == indKey) != null)
                //    break;
                if (IsReadableRead(first))
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

            if (last != null)
            {
                int firstRow = last.RowNum;
                int lastrow = workSheet.LastRowNum;

                for (int index = firstRow; index <= lastrow; index++)
                {
                    IRow tmp = workSheet.GetRow(index);
                    if (!IsReadableRead(tmp))
                        break;
                    last = tmp;
                }
            }
            return last;
        }
    }
}

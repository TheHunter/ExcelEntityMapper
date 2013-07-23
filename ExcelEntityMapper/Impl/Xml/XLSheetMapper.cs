using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using ExcelEntityMapper.Exceptions;

namespace ExcelEntityMapper.Impl.Xml
{
    /// <summary>
    /// 
    /// </summary>
    /// <typeparam name="TSource"></typeparam>
    public class XLSheetMapper<TSource>
        : SheetMapper<TSource>, IXLSheetWorker<TSource>, IXLWorkBookReader<TSource>, IXLWorkBookWriter<TSource>
        where TSource : class, new()
    {
        private XLWorkbook workBook;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="propertyMappers"></param>
        public XLSheetMapper(IEnumerable<IXLPropertyMapper<TSource>> propertyMappers)
            : this(0, propertyMappers)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="headerRows"></param>
        /// <param name="propertyMappers"></param>
        public XLSheetMapper(int headerRows, IEnumerable<IXLPropertyMapper<TSource>> propertyMappers)
            : base(headerRows, false, propertyMappers)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        XLWorkbook IXLWorkBookProvider.WorkBook
        {
            get { return this.workBook; }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="index"></param>
        /// <param name="color"></param>
        public void ChangeBackGround(int index, XLColor color)
        {
            IXLWorksheet workSheet = this.GetWorkSheet(this.SheetName);
            workSheet.Row(index)
                .Cells(this.FirstColumn, this.LastColumn)
                .Style
                .Fill
                .BackgroundColor = color;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="rowBase"></param>
        /// <param name="sheet"></param>
        protected void WriteHeader(int rowBase, IXLWorksheet sheet)
        {
            var row = sheet.Row(rowBase);
            
            this.PropertyMappers.All
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
        /// <param name="workbook"></param>
        public void InjectWorkBook(IXLWorkBook workbook)
        {
            if (workbook == null)
                throw new WorkBookException("Impossible to associate a null workbook.");

            XLWorkBook current = workbook as XLWorkBook;
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

            IXLWorksheet workSheet = this.GetWorkSheet(sheetName);
            IXLRow row = this.GetFirstRow(workSheet);

            if (row == null)
                return 0;

            if (this.HasHeader)
                row = row.RowBelow(this.HeaderRows);

            int counter = 0;
            while (IsReadableRow(row))
            {
                counter++;
                try
                {
                    TSource current = this.ReadInstance(row);
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
        /// <param name="row"></param>
        /// <returns></returns>
        private TSource ReadInstance(IXLRow row)
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
        private bool IsReadableRow(IXLRow row)
        {
            var col = this.PropertyMappers.Where(n => n.CustomType == MapperType.Key);
            return col.Any() && col.All(n => !row.Cell(n.ColumnIndex).IsEmpty());
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

            IXLWorksheet workSheet = this.GetWorkSheet(sheetName);

            int counter = 0;
            if (instances != null && instances.Any())
            {
                int rowIndex = 0;
                //IXLCell header = workSheet.Column(this.IndexKeyColumn).LastCellUsed(); //
                IXLRow lastRow = this.GetLastRow(workSheet);

                if (this.HasHeader)
                {
                    if (lastRow == null)
                    {
                        rowIndex = 1;
                        this.WriteHeader(rowIndex, workSheet);
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
                    if (this.WriteInstance(row, current))
                        counter++;
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
        private bool WriteInstance(IXLRow row, TSource instance)
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

            if (this.AfterWriting != null)
                this.AfterWriting(instance);

            return ret;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheetname"></param>
        /// <returns></returns>
        private IXLWorksheet GetWorkSheet(string sheetname)
        {
            try
            {
                return this.workBook.Worksheet(sheetname);
            }
            catch (Exception ex)
            {
                throw new NotAvailableWorkSheetException("Impossible to find a worksheet with the specified name.", sheetname, ex);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="workSheet"></param>
        /// <returns></returns>
        private IXLRow GetFirstRow(IXLWorksheet workSheet)
        {
            IXLRow firstRow = null;
            IXLRow lastRow = workSheet.LastRowUsed();

            if (lastRow != null)
            {
                int firstIndex = 1;
                int lastIndex = lastRow.RowNumber();

                for (int index = firstIndex; index <= lastIndex; index++)
                {
                    lastRow = workSheet.Row(index);
                    if (IsReadableRow(lastRow))
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
        /// <param name="workSheet"></param>
        /// <returns></returns>
        private IXLRow GetLastRow(IXLWorksheet workSheet)
        {
            IXLRow firstRow = GetFirstRow(workSheet);
            IXLRow lastRow = workSheet.LastRowUsed();

            if (firstRow != null && lastRow != null)
            {
                int firstIndex = firstRow.RowNumber();
                int lastIndex = lastRow.RowNumber();

                for (int index = firstIndex; index <= lastIndex; index++)
                {
                    IXLRow currRow = workSheet.Row(index);
                    if (!IsReadableRow(currRow))
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
    }
}

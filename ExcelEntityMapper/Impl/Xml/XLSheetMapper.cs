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
        : SheetMapper<TSource>, IXLSheetWorker<TSource>
        where TSource : class, new()
    {
        private XLWorkbook workBook;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="indexkeyColumn"></param>
        /// <param name="propertyMappers"></param>
        public XLSheetMapper(int indexkeyColumn, IEnumerable<IXLPropertyMapper<TSource>> propertyMappers)
            : this(indexkeyColumn, false, propertyMappers)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="indexkeyColumn"></param>
        /// <param name="hasHeader"></param>
        /// <param name="propertyMappers"></param>
        public XLSheetMapper(int indexkeyColumn, bool hasHeader, IEnumerable<IXLPropertyMapper<TSource>> propertyMappers)
            : base(indexkeyColumn, hasHeader, false, propertyMappers)
        {
            this.LastColumn = this.PropertyMappers.Select(n => n.ColumnIndex).Max();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="index"></param>
        /// <param name="color"></param>
        public void ChangeBackGround(int index, XLColor color)
        {
            IXLWorksheet workSheet = this.GetWorkSheet(this.SheetName);
            workSheet.Row(index).Cells(this.IndexKeyColumn, this.LastColumn).Style.Fill.BackgroundColor = color;
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

            IXLWorksheet workSheet = this.GetWorkSheet(sheetName);
            IXLRow row = this.GetFirstRow(workSheet);

            if (row == null)
                return 0;

            if (this.HasHeader)
                row = row.RowBelow();

            int counter = 0;
            while (!row.Cell(this.IndexKeyColumn).IsEmpty())
            {
                counter++;
                TSource current = this.ReadInstance(row);
                if (current != null)
                {
                    buffer.Add(row.RowNumber(), current);
                }
                row = row.RowBelow();
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
                IXLCell cell;
                this.PropertyMappers.All
                    (
                        parameter =>
                            {
                                try
                                {
                                    cell = row.Cell(parameter.ColumnIndex);
                                    parameter.ToPropertyFormat(instance, cell.GetString().Trim());
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

            IXLWorksheet workSheet = this.GetWorkSheet(sheetName);

            int counter = 0;
            if (instances != null && instances.Any())
            {
                IXLRow row;
                int rowIndex = 0;
                IXLCell header = workSheet.Column(this.IndexKeyColumn).LastCellUsed();

                if (this.HasHeader)
                {
                    if (header == null)
                    {
                        rowIndex = 1;
                        this.WriteHeader(rowIndex, workSheet);
                    }
                    else
                    {
                        rowIndex = header.Address.RowNumber;
                    }
                }

                foreach (var current in instances)
                {
                    rowIndex++;
                    row = workSheet.Row(rowIndex);
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
        private bool WriteInstance(IXLRow row, TSource instance)
        {
            bool ret = false;
            if (instance != null && row != null)
            {
                IXLCell cell = null;
                this.PropertyMappers.All
                (
                    parameter =>
                        {
                            try
                            {
                                cell = row.Cell(parameter.ColumnIndex);
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
            IXLCell current = workSheet.Column(this.IndexKeyColumn).FirstCellUsed();
            if (current != null)
                return current.WorksheetRow();
            return null;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="workSheet"></param>
        /// <returns></returns>
        private IXLRow GetLastRow(IXLWorksheet workSheet)
        {
            IXLCell current = workSheet.Column(this.IndexKeyColumn).LastCellUsed();
            if (current != null)
                return current.WorksheetRow();
            return null;
        }
    }
}

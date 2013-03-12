using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SmartXLS;
using System.Reflection;

namespace ExcelEntityMapper.Impl
{
    /// <summary>
    /// 
    /// </summary>
    /// <typeparam name="TSource"></typeparam>
    public class XSheetMapper<TSource>
        : XSheetBase, IXLSheetWorker<TSource>
        where TSource : class, new()
    {
        private readonly List<IXLPropertyMapper<TSource>> _Parameters = null;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="indexkeyColumn"></param>
        /// <param name="parameters"></param>
        public XSheetMapper(int indexkeyColumn, IEnumerable<IXLPropertyMapper<TSource>> parameters)
            : this(indexkeyColumn, false, parameters)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="indexkeyColumn"></param>
        /// <param name="hasHeader"></param>
        /// <param name="parameters"></param>
        public XSheetMapper(int indexkeyColumn, bool hasHeader, IEnumerable<IXLPropertyMapper<TSource>> parameters)
            : base(indexkeyColumn, hasHeader, true)
        {
            if (parameters == null || parameters.Count() == 0)
                throw new ArgumentNullException("Column parameters cannot be null or empty.");

            this._Parameters = new List<IXLPropertyMapper<TSource>>(parameters);
            this.LastColumn = this._Parameters.Select(n => n.ColumnIndex).Max() - this.Offset;
        }

        //internal WorkBook XLWorkbook
        //{
        //    get { return this._XLWorkbook; }
        //    set { this._XLWorkbook = value; }
        //}

        /// <summary>
        /// 
        /// </summary>
        public IEnumerable<IXLPropertyMapper<TSource>> PropertyMappers
        {
            get { return _Parameters; }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="buffer"></param>
        /// <returns></returns>
        public int ReadObjects(IXLWorkBook workbook, IDictionary<int, TSource> buffer)
        {
            return this.ReadObjects(workbook, this.SheetName, buffer);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="sheetName"></param>
        /// <param name="buffer"></param>
        /// <returns></returns>
        public int ReadObjects(IXLWorkBook workbook, string sheetName, IDictionary<int, TSource> buffer)
        {
            TSource current = null;
            int counter = 0;

            WorkBook wb = SourceHelper.VerifyWorkBook(workbook, sheetName);
            int? currentRow = GetFirstRow(wb);

            if (currentRow.HasValue)
            {
                if (this.HasHeader)
                {
                    currentRow = currentRow.Value + 1;
                }

                int indexCol = this.IndexKeyColumn - this.Offset;
                while (!string.IsNullOrEmpty(wb.getText(currentRow.Value, indexCol)))
                {
                    counter++;
                    current = this.ReadInstance(wb, currentRow.Value);
                    if (current != null)
                    {
                        buffer.Add(currentRow.Value + this.Offset, current);
                    }
                    currentRow = currentRow.Value + 1;
                }
            }

            return counter;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private int? GetFirstRow(WorkBook workbook)
        {
            int? ret = null;
            int index = 0;
            int last = workbook.LastRow;
            int indexCol = this.IndexKeyColumn - this.Offset;

            while (index <= last)
            {
                if (!string.IsNullOrEmpty(workbook.getText(index, indexCol)))
                {
                    ret = index;
                    break;
                }
                index++;
            }
            return ret;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private int? GetLastRow(WorkBook workbook)
        {
            int? ret = null;
            int index = workbook.LastRow;
            int rowbase = this.ZeroBase ? 0 : 1;
            int indexCol = this.IndexKeyColumn - this.Offset;

            while (index >= rowbase)
            {
                if (!string.IsNullOrEmpty(workbook.getText(index, indexCol)))
                {
                    ret = index;
                    break;
                }
                index--;
            }
            return ret;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        private TSource ReadInstance(WorkBook workbook, int row)
        {
            TSource instance = null;
            instance = SourceHelper.CreateInstance<TSource>();

            this.PropertyMappers.All
                (
                    delegate(IXLPropertyMapper<TSource> parameter)
                    {
                        try
                        {
                            parameter.ToPropertyFormat(instance, workbook.getText(row, parameter.ColumnIndex - this.Offset));
                            //parameter.ToPropertyFormat(instance, workbook.getText(row, parameter.ColumnIndex - this.Offset, false));
                            //parameter.ToPropertyFormat(instance, workbook.getFormattedText(row, parameter.ColumnIndex - this.Offset));
                        }
                        catch (Exception)
                        {
                            instance = null;
                            return false;
                        }
                        return true;
                    }
                );

            return instance;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="instances"></param>
        /// <returns></returns>
        public int WriteObjects(IXLWorkBook workbook, IEnumerable<TSource> instances)
        {
            return this.WriteObjects(workbook, this.SheetName, instances);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="sheetName"></param>
        /// <param name="instances"></param>
        /// <returns></returns>
        public int WriteObjects(IXLWorkBook workbook, string sheetName, IEnumerable<TSource> instances)
        {
            WorkBook wb = SourceHelper.VerifyWorkBook(workbook, sheetName);

            int counter = 0;

            if (instances != null && instances.Count() > 0)
            {
                int rowbase = this.ZeroBase ? 0 : 1;
                int? last = this.GetLastRow(wb);
                int row = last.HasValue ? last.Value : rowbase;

                if (!last.HasValue && this.HasHeader)
                {
                    this.WriteHeader(wb, rowbase);
                    row++;
                }

                foreach (var current in instances)
                {
                    if (this.WriteInstance(row, wb, current))
                    {
                        counter++;
                    }
                    row++;
                }
            }

            return counter;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="rowBase"></param>
        private void WriteHeader(WorkBook workbook, int rowBase)
        {
            this.PropertyMappers.All
                (
                    delegate(IXLPropertyMapper<TSource> parameter)
                    {
                        workbook.setText(rowBase, parameter.ColumnIndex - this.Offset, parameter.ColumnHeader);
                        return true;
                    }
                );
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="row"></param>
        /// <param name="workbook"></param>
        /// <param name="instance"></param>
        /// <returns></returns>
        private bool WriteInstance(int row, WorkBook workbook, TSource instance)
        {
            bool ret = false;
            if (instance != null)
            {
                this._Parameters.All
                (
                    delegate(IXLPropertyMapper<TSource> parameter)
                    {
                        try
                        {
                            workbook.setText(row, parameter.ColumnIndex - this.Offset, parameter.ToExcelFormat(instance));
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

        
        //public bool InjectWorkSheet(IXLWorkBook workbook)
        //{
        //    return this.InjectWorkSheet(workbook, this.SheetName);
        //}

        
        //public bool InjectWorkSheet(IXLWorkBook workbook, string sheetname)
        //{
        //    bool ret = false;

        //    if (workbook == null)
        //        throw new ArgumentNullException("workbook non puo' essere null.");

        //    if (string.IsNullOrEmpty(sheetname))
        //        throw new ArgumentException("sheetname non puo' essere null oppure vuoto.");

        //    XWorkBook temp = workbook as XWorkBook;
        //    if (temp != null)
        //    {
        //        this._XLWorkbook = temp.Workbook;
        //        //this._SheetIndex = this._XLWorkbook.findSheetByName(sheetname);
        //    }
        //    else
        //    {
        //        throw new InvalidCastException("workbook type is not valid.");
        //    }

        //    return ret;
        //}
    }
}

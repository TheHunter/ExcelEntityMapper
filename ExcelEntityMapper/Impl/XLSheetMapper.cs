using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

namespace ExcelEntityMapper.Impl
{
    /// <summary>
    /// 
    /// </summary>
    /// <typeparam name="TSource"></typeparam>
    internal class XLSheetMapper<TSource>
        : XLSheetBase //, IXLSheetWorker<TSource>
        where TSource : class, new()
    {
        private IXLWorksheet _WorkSheet = null;
        private readonly List<IXLPropertyMapper<TSource>> _Parameters = null;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="indexkeyColumn"></param>
        /// <param name="parameters"></param>
        public XLSheetMapper(int indexkeyColumn, IEnumerable<IXLPropertyMapper<TSource>> parameters)
            :this(indexkeyColumn, false, parameters)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="indexkeyColumn"></param>
        /// <param name="hasHeader"></param>
        /// <param name="parameters"></param>
        public XLSheetMapper(int indexkeyColumn, bool hasHeader, IEnumerable<IXLPropertyMapper<TSource>> parameters)
            : base(indexkeyColumn, hasHeader)
        {
            if (parameters == null || parameters.Count() == 0)
                throw new ArgumentNullException("Column parameters cannot be null or empty.");

            this._Parameters = new List<IXLPropertyMapper<TSource>>(parameters);
            this.LastColumn = this._Parameters.Select(n => n.ColumnIndex).Max();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="index"></param>
        /// <param name="color"></param>
        protected override void ChangeBackGround(int index, ClosedXML.Excel.XLColor color)
        {
            this.WorkSheet.Row(index).Cells(this.IndexKeyColumn, this.LastColumn).Style.Fill.BackgroundColor = color;
        }

        /// <summary>
        /// 
        /// </summary>
        public IEnumerable<IXLPropertyMapper<TSource>> Parameters
        {
            get { return this._Parameters; }
        }

        /// <summary>
        /// 
        /// </summary>
        internal IXLWorksheet WorkSheet
        {
            get { return this._WorkSheet; }
            private set { this._WorkSheet = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="buffer"></param>
        /// <returns></returns>
        public int ReadObjects(IDictionary<int, TSource> buffer)
        {
            if (this.WorkSheet == null)
                throw new NotImplementedException("Nessun worksheet associato al'oggetto chiamante, assegnare un worksheet tramite il metodo InjectWorkSheet() ");

            IXLRow row = this.WorkSheet.FirstRowUsed();
            TSource current = null;
            int counter = 0;

            if (this.HasHeader)
            {
                row = row.RowBelow();
            }

            while (!row.Cell(this.IndexKeyColumn).IsEmpty())
            {
                counter++;
                current = this.ReadInstance(row);
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
                IXLCell cell = null;
                this.Parameters.All
                    (
                        delegate(IXLPropertyMapper<TSource> parameter)
                        {
                            try
                            {
                                cell = row.Cell(parameter.ColumnIndex);
                                parameter.ToPropertyFormat(instance, cell.GetString().Trim());
                            }
                            finally
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
            if (this.WorkSheet == null)
                throw new NotImplementedException("Nessun worksheet associato al'oggetto chiamante, assegnare un worksheet tramite il metodo InjectWorkSheet() ");

            int counter = 0;
            if (instances != null && instances.Count() > 0)
            {
                IXLRow row = this.WorkSheet.LastRowUsed();
                foreach (var current in instances)
                {
                    row = row.RowBelow();
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
                this._Parameters.All
                (
                    delegate(IXLPropertyMapper<TSource> parameter)
                    {
                        try
                        {
                            cell = row.Cell(parameter.ColumnIndex);
                            cell.Value = parameter.ToExcelFormat(instance);
                        }
                        finally
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
        /// <param name="workbook"></param>
        /// <returns></returns>
        public bool InjectWorkSheet(IXLWorkBook workbook)
        {
            return this.InjectWorkSheet(workbook, this.SheetName);
        }

        /// <summary>
        /// Associa il worksheet indicato nell'argomento "sheetname" presente nel workbook.
        /// </summary>
        /// <param name="workbook">Oggetto workbook che contiene tutti i fogli di excel.</param>
        /// <param name="sheetname">WorkSheet che contiene i dati da leggere / scrivere.</param>
        /// <returns>Se il worksheet viene trovato allora torna true, altrimenti torna false.</returns>
        public bool InjectWorkSheet(IXLWorkBook workbook, string sheetname)
        {
            bool ret = false;

            if (workbook == null)
                throw new ArgumentNullException("workbook non puo' essere null.");

            if (string.IsNullOrEmpty(sheetname))
                throw new ArgumentException("sheetname non puo' essere null oppure vuoto.");

            XLWorkBook temp = workbook as XLWorkBook;
            if (temp != null)
            {
                var sheet = temp.Workbook.Worksheet(sheetname);
                if (sheet != null)
                {
                    ret = true;
                    this.WorkSheet = sheet;
                }
            }
            else
            {
                throw new InvalidCastException("workbook type is not valid.");
            }

            return ret;
        }
    }
}

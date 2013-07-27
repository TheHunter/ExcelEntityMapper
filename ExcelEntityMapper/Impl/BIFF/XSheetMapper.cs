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
        : SheetMapper<TSource>, IXLSheetWorker<TSource>, IXWorkBookReader<TSource>, IXWorkBookWriter<TSource>
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
        }

        /// <summary>
        /// 
        /// </summary>
        HSSFWorkbook IXWorkBookProvider<TSource>.WorkBook
        {
            get { return workBook; }
            set { this.workBook = value; }
        }
        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="workbook"></param>
        public override void InjectWorkBook(IXLWorkBook workbook)
        {
            this.InjectWorkBook<TSource>(workbook);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="buffer"></param>
        /// <returns></returns>
        public override int ReadObjects(string sheetName, IDictionary<int, TSource> buffer)
        {
            return this.ReadObjects<TSource>(sheetName, buffer);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="instances"></param>
        /// <returns></returns>
        public override int WriteObjects(string sheetName, IEnumerable<TSource> instances)
        {
            return this.WriteObjects<TSource>(sheetName, instances);
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.HSSF.UserModel;

namespace ExcelEntityMapper.Impl.BIFF
{
    /// <summary>
    /// 
    /// </summary>
    /// <typeparam name="TSource"></typeparam>
    public class XSheetReader<TSource>
        : SheetReader<TSource>, IXWorkBookReader<TSource>
        where TSource : class, new()
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="propertyMappers"></param>
        public XSheetReader(IEnumerable<IXLPropertyMapper<TSource>> propertyMappers)
            : this(0, propertyMappers)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="headerRows"></param>
        /// <param name="propertyMappers"></param>
        public XSheetReader(int headerRows, IEnumerable<IXLPropertyMapper<TSource>> propertyMappers)
            : base(headerRows, true, propertyMappers)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        HSSFWorkbook IXWorkBookProvider<TSource>.WorkBook { get; set; }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public override int GetIndexFirstRow(string sheetName)
        {
            return this.GetIndexFirstRow<TSource>(sheetName);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="rowIndex"></param>
        /// <returns></returns>
        public override TSource ReadObject(string sheetName, int rowIndex)
        {
            return this.ReadObject<TSource>(sheetName, rowIndex);
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
        /// <param name="workbook"></param>
        public override void InjectWorkBook(IXLWorkBook workbook)
        {
            this.InjectWorkBook<TSource>(workbook);
        }
    }
}

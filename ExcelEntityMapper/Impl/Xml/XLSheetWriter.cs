using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

namespace ExcelEntityMapper.Impl.Xml
{
    /// <summary>
    /// 
    /// </summary>
    public class XLSheetWriter<TSource>
        : SheetWriter<TSource>, IXLWorkBookWriter<TSource>
        where TSource : class
    {
        private XLWorkbook workBook;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="headerRows"></param>
        /// <param name="zeroBase"></param>
        /// <param name="propertyMappers"></param>
        public XLSheetWriter(int headerRows, bool zeroBase, IEnumerable<IXLPropertyMapper<TSource>> propertyMappers)
            : base(headerRows, zeroBase, propertyMappers)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        XLWorkbook IXLWorkBookProvider<TSource>.WorkBook { get { return this.workBook; } }

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

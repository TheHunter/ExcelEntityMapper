using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelEntityMapper.Impl.Xml
{
    /// <summary>
    /// 
    /// </summary>
    /// <typeparam name="TSource"></typeparam>
    public class XLSheetReader<TSource>
        : SheetReader<TSource>
        where TSource : class
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="headerRows"></param>
        /// <param name="zeroBase"></param>
        /// <param name="propertyMappers"></param>
        public XLSheetReader(int headerRows, bool zeroBase, IEnumerable<IXLPropertyMapper<TSource>> propertyMappers)
            : base(headerRows, zeroBase, propertyMappers)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="buffer"></param>
        /// <returns></returns>
        public override int ReadObjects(string sheetName, IDictionary<int, TSource> buffer)
        {
            throw new NotImplementedException();
        }
    }
}

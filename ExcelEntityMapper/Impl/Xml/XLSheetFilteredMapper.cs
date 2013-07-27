using System;
using System.Collections.Generic;
using System.Linq;
using ExcelEntityMapper.Exceptions;

namespace ExcelEntityMapper.Impl.Xml
{
    /// <summary>
    /// 
    /// </summary>
    /// <typeparam name="TSource"></typeparam>
    public class XLSheetFilteredMapper<TSource>
        : XLSheetMapper<TSource>, IXLSheetFiltered<TSource>
        where TSource : class, new()
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="parameters"></param>
        public XLSheetFilteredMapper(IEnumerable<IXLPropertyMapper<TSource>> parameters)
            : this(0, parameters)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="headerRows"></param>
        /// <param name="parameters"></param>
        public XLSheetFilteredMapper(int headerRows, IEnumerable<IXLPropertyMapper<TSource>> parameters)
            : base(headerRows, parameters)
        {

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="buffer"></param>
        /// <param name="function"></param>
        /// <returns></returns>
        public int ReadFilteredObjects(IDictionary<int, TSource> buffer, Func<TSource, bool> function)
        {
            return ReadFilteredObjects(this.SheetName, buffer, function);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="buffer"></param>
        /// <param name="function"></param>
        /// <returns></returns>
        public int ReadFilteredObjects(string sheetName, IDictionary<int, TSource> buffer, Func<TSource, bool> function)
        {
            return this.ReadFilteredObjects<TSource>(sheetName, buffer, function);
        }
    }
}

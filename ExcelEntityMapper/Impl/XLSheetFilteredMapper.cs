using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelEntityMapper.Impl
{
    /// <summary>
    /// 
    /// </summary>
    /// <typeparam name="TSource"></typeparam>
    internal class XLSheetFilteredMapper<TSource>
        : XLSheetMapper<TSource>//, IXLSheetFiltered<TSource>
        where TSource : class, new()
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="indexkeyColumn"></param>
        /// <param name="parameters"></param>
        public XLSheetFilteredMapper(int indexkeyColumn, IEnumerable<IXLPropertyMapper<TSource>> parameters)
            : this(indexkeyColumn, false, parameters)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="indexkeyColumn"></param>
        /// <param name="hasHeader"></param>
        /// <param name="parameters"></param>
        public XLSheetFilteredMapper(int indexkeyColumn, bool hasHeader, IEnumerable<IXLPropertyMapper<TSource>> parameters)
            : base(indexkeyColumn, hasHeader, parameters)
        {

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="function"></param>
        /// <param name="buffer"></param>
        /// <returns></returns>
        public int ReadFilteredObjects(Func<TSource, bool> function, IDictionary<int, TSource> buffer)
        {
            throw new NotImplementedException();
        }
    }
}

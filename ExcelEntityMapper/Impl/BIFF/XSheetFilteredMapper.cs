using System;
using System.Collections.Generic;
using System.Linq;
using ExcelEntityMapper.Exceptions;

namespace ExcelEntityMapper.Impl.BIFF
{
    /// <summary>
    /// 
    /// </summary>
    /// <typeparam name="TSource"></typeparam>
    public class XSheetFilteredMapper<TSource>
        : XSheetMapper<TSource>, IXLSheetFiltered<TSource>
        where TSource : class, new()
    {

        /// <summary>
        /// 
        /// </summary>
        /// <param name="indexkeyColumn"></param>
        /// <param name="parameters"></param>
        public XSheetFilteredMapper(int indexkeyColumn, IEnumerable<IXLPropertyMapper<TSource>> parameters)
            : this(indexkeyColumn, 0, parameters)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="indexkeyColumn"></param>
        /// <param name="headerRows"></param>
        /// <param name="parameters"></param>
        public XSheetFilteredMapper(int indexkeyColumn, int headerRows, IEnumerable<IXLPropertyMapper<TSource>> parameters)
            : base(indexkeyColumn, headerRows, parameters)
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
            return this.ReadFilteredObjects(this.SheetName, buffer, function);
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
            if (buffer == null)
                throw new SheetParameterException("Buffer cannot be null.", "buffer");

            Dictionary<int, TSource> temp = new Dictionary<int, TSource>();
            int count = this.ReadObjects(temp);

            var filtered = temp.Where(n => function(n.Value));

            if (filtered.Any())
            {
                foreach (var current in filtered)
                {
                    buffer.Add(current);
                }
            }

            return count;
        }
    }
}

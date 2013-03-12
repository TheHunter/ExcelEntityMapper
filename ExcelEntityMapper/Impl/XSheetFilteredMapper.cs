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
    public class XSheetFilteredMapper<TSource>
        : XSheetMapper<TSource>, IXLSheetFiltered<TSource>
        where TSource : class, new()
    {
        private IDictionary<int, TSource> _Buffer = new Dictionary<int, TSource>();

        /// <summary>
        /// 
        /// </summary>
        /// <param name="indexkeyColumn"></param>
        /// <param name="parameters"></param>
        public XSheetFilteredMapper(int indexkeyColumn, IEnumerable<IXLPropertyMapper<TSource>> parameters)
            : this(indexkeyColumn, false, parameters)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="indexkeyColumn"></param>
        /// <param name="hasHeader"></param>
        /// <param name="parameters"></param>
        public XSheetFilteredMapper(int indexkeyColumn, bool hasHeader, IEnumerable<IXLPropertyMapper<TSource>> parameters)
            : base(indexkeyColumn, hasHeader, parameters)
        {

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="buffer"></param>
        /// <param name="function"></param>
        /// <returns></returns>
        public int ReadFilteredObjects(IXLWorkBook workbook, IDictionary<int, TSource> buffer, Func<TSource, bool> function)
        {
            return this.ReadFilteredObjects(workbook, this.SheetName, buffer, function);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="sheetName"></param>
        /// <param name="buffer"></param>
        /// <param name="function"></param>
        /// <returns></returns>
        public int ReadFilteredObjects(IXLWorkBook workbook, string sheetName, IDictionary<int, TSource> buffer, Func<TSource, bool> function)
        {
            if (buffer == null)
                throw new ArgumentNullException("Buffer cannot be null.");

            _Buffer.Clear();
            int count = this.ReadObjects(workbook, _Buffer);

            var filtered = _Buffer.Where(n => function(n.Value));

            if (filtered != null)
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

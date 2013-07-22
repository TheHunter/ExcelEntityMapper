using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.HSSF.UserModel;

namespace ExcelEntityMapper
{
    /// <summary>
    /// 
    /// </summary>
    interface IXWorkBookProvider<TSource>
        where TSource : class
    {
        /// <summary>
        /// 
        /// </summary>
        HSSFWorkbook WorkBook { get; }

        IEnumerable<IXLPropertyMapper<TSource>> PropertyMappers { get; };
    }
}

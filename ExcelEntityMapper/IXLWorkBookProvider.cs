using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using NPOI.HSSF.UserModel;

namespace ExcelEntityMapper
{

    /// <summary>
    /// 
    /// </summary>
    interface IXWorkBookProvider<TSource>
        : IXLSheetMapper<TSource>
        where TSource : class
    {
        /// <summary>
        /// 
        /// </summary>
        HSSFWorkbook WorkBook { get; }
    }

    /// <summary>
    /// 
    /// </summary>
    internal interface IXLWorkBookProvider<TSource>
        : IXLSheetMapper<TSource>
        where TSource : class
    {
        /// <summary>
        /// 
        /// </summary>
        XLWorkbook WorkBook { get; }
    }

    
}

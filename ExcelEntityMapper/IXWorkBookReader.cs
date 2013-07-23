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
    interface IXWorkBookReader<TSource>
        : IXLSheetReader<TSource>, IXWorkBookProvider<TSource>
        where TSource : class, new()
    {
    }

    /// <summary>
    /// 
    /// </summary>
    /// <typeparam name="TSource"></typeparam>
    interface IXLWorkBookReader<TSource>
        : IXLSheetReader<TSource>, IXLWorkBookProvider<TSource>
        where TSource : class, new()
    {
    }
}

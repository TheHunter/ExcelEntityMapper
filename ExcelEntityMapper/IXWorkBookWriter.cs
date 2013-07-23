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
    /// <typeparam name="TSource"></typeparam>
    interface IXWorkBookWriter<TSource>
        : IXLSheetWriter<TSource>, IXWorkBookProvider
        where TSource : class
    {
    }

    /// <summary>
    /// 
    /// </summary>
    /// <typeparam name="TSource"></typeparam>
    interface IXLWorkBookWriter<TSource>
        : IXLSheetWriter<TSource>, IXLWorkBookProvider
        where TSource : class
    {
    }
}

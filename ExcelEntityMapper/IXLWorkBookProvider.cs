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
    interface IXWorkBookProvider
    {
        /// <summary>
        /// 
        /// </summary>
        HSSFWorkbook WorkBook { get; }
    }

    /// <summary>
    /// 
    /// </summary>
    public interface IXLWorkBookProvider
    {
        /// <summary>
        /// 
        /// </summary>
        XLWorkbook WorkBook { get; }
    }

    
}

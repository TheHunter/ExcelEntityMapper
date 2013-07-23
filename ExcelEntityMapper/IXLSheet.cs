using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelEntityMapper
{
    /// <summary>
    /// Defines a default members for standars worksheets.
    /// </summary>
    public interface IXLSheet
    {
        /// <summary>
        /// The sheetname to use for reading / writting instances.
        /// </summary>
        string SheetName { get; set; }

        /// <summary>
        /// Rappresents the first column mapped.
        /// </summary>
        int FirstColumn { get; }

        /// <summary>
        /// Rappresents the last column mapped.
        /// </summary>
        int LastColumn { get; }

        /// <summary>
        /// Indicates if exists a header.
        /// </summary>
        bool HasHeader { get; }

        /// <summary>
        /// Indicates the numbr of rows which is composed the header.
        /// </summary>
        int HeaderRows { get; }

        /// <summary>
        /// 
        /// </summary>
        bool ZeroBase { get; }

        /// <summary>
        /// 
        /// </summary>
        int Offset { get; }
    }
}

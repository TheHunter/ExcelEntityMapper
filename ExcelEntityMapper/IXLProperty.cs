using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelEntityMapper
{
    /// <summary>
    /// 
    /// </summary>
    public interface IXLProperty
    {
        /// <summary>
        /// Indicates the column index for map property typed class.
        /// </summary>
        int ColumnIndex { get; }

        /// <summary>
        /// Rappresents the default header name for this mapper.
        /// </summary>
        string ColumnHeader { get; }

        /// <summary>
        /// Rapresents the mapper type for this mapper.
        /// </summary>
        MapperType CustomType { get; }

        /// <summary>
        /// Indicates the operations that can be done with the source property.
        /// </summary>
        SourceOperation OperationEnabled { get; }
    }
}

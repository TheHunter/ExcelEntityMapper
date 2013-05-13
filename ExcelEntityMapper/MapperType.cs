using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelEntityMapper
{
    /// <summary>
    /// 
    /// </summary>
    public enum MapperType
    {
        /// <summary>
        /// Indicates that the PropertyMapper value must be found.
        /// </summary>
        Required,

        /// <summary>
        /// Indicates that this kind of mapper is required, and It can be used as a Key for searching valid rows.
        /// </summary>
        Key,
        
        /// <summary>
        /// A simple mapper that's no required.
        /// </summary>
        Simple
    }
}

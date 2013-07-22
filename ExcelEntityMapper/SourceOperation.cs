using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelEntityMapper
{
    /// <summary>
    /// Indicates the kind of operation can be done by property mapper instances.
    /// </summary>
    public enum SourceOperation
    {
        /// <summary>
        /// Indicates the property mapper is read only.
        /// </summary>
        Read,
        
        /// <summary>
        /// Indicates the property mapper can be written by Excel cell value.
        /// </summary>
        Write,

        /// <summary>
        /// Indicates the property mapper is bidirectional, so It's writeable and readable.
        /// </summary>
        ReadWrite
    }
}

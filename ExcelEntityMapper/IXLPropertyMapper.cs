using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelEntityMapper
{
    /// <summary>
    /// Defines a set of properties for customizing property mappers.
    /// </summary>
    /// <typeparam name="TSource">Type of instance which property mapper is used.</typeparam>
    public interface IXLPropertyMapper<TSource>
        where TSource : class
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
        /// The action which serves for setting the property instance.
        /// </summary>
        Action<TSource, string> ToPropertyFormat { get; }

        /// <summary>
        /// A function which serves for transform cell value into property type value.
        /// </summary>
        Func<TSource, string> ToExcelFormat { get; }
    }
}

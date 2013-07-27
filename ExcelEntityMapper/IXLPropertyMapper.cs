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
        : IXLPropertyWriter<TSource>, IXLPropertyReader<TSource>
        where TSource : class
    {
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using ExcelEntityMapper.Exceptions;
using ExcelEntityMapper.Impl;
using NPOI.SS.UserModel;

namespace ExcelEntityMapper
{
    public class FactoryMapper
    {
        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="column"></param>
        /// <param name="mapperType"></param>
        /// <param name="columnHeader"></param>
        /// <param name="toExcelFormat"></param>
        /// <returns></returns>
        public static IXLPropertyMapper<TSource> MakeReaderPropertyMap<TSource>(int column, MapperType mapperType, string columnHeader, Expression<Func<TSource, string>> toExcelFormat)
            where TSource : class
        {
            return new PropertyMapper<TSource>(column, mapperType, columnHeader, toExcelFormat);
        }

        
    }
}

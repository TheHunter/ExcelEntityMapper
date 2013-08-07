using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using ClosedXML.Excel;
using ExcelEntityMapper.Exceptions;
using ExcelEntityMapper.Impl;
using ExcelEntityMapper.Impl.BIFF;
using ExcelEntityMapper.Impl.Xml;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace ExcelEntityMapper
{
    /// <summary>
    /// 
    /// </summary>
    public static class FactoryMapper
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
        public static IXLPropertyMapper<TSource> MakeReaderPropertyMap<TSource>(int column, MapperType mapperType, string columnHeader, Expression<Func<TSource, object>> toExcelFormat)
            where TSource : class
        {
            return new PropertyMapper<TSource>(column, mapperType, columnHeader, toExcelFormat);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="column"></param>
        /// <param name="mapperType"></param>
        /// <param name="columnHeader"></param>
        /// <param name="toPropertyFormat"></param>
        /// <returns></returns>
        public static IXLPropertyMapper<TSource> MakeWriterPropertyMap<TSource>(int column, MapperType mapperType, string columnHeader, Action<TSource, string> toPropertyFormat)
            where TSource : class
        {
            return new PropertyMapper<TSource>(column, mapperType, columnHeader, toPropertyFormat);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="inputStream"></param>
        /// <returns></returns>
        public static IXLWorkBook MakeWorkBook(byte[] inputStream)
        {
            return FactoryMapper.ResolveWorkBook(inputStream);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="inputStream"></param>
        /// <returns></returns>
        public static IXLWorkBook MakeWorkBook(Stream inputStream)
        {
            byte[] input = new byte[inputStream.Length];
            inputStream.Read(input, 0, input.Length);
            return FactoryMapper.ResolveWorkBook(input);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="inputStream"></param>
        /// <returns></returns>
        private static IXLWorkBook ResolveWorkBook(byte[] inputStream)
        {
            try
            {
                using (MemoryStream mem = new MemoryStream(inputStream))
                    return new XWorkBook(new HSSFWorkbook(mem));
            }
            catch (Exception)
            {
                try
                {
                    return new XLWorkBook(inputStream);
                }
                catch (Exception ex)
                {
                    throw new WorkBookException("Error on reading the stream input.", ex);
                }
            }
        }
    }
}

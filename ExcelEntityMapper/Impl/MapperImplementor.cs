using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using ClosedXML.Excel;
using ExcelEntityMapper.Exceptions;
using ExcelEntityMapper.Impl.BIFF;
using ExcelEntityMapper.Impl.Xml;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace ExcelEntityMapper.Impl
{
    /// <summary>
    /// Extension class for sheet mappers.
    /// </summary>
    public static class MapperImplementor
    {
        


        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="sheetMapper"></param>
        /// <param name="sheetName"></param>
        /// <param name="buffer"></param>
        /// <param name="function"></param>
        /// <returns></returns>
        internal static int ReadFilteredObjects<TSource>(this IXLSheetFiltered<TSource> sheetMapper, string sheetName, IDictionary<int, TSource> buffer,
                                                Func<TSource, bool> function)
             where TSource : class, new()
        {
            if (buffer == null)
                throw new SheetParameterException("Buffer cannot be null.", "buffer");

            Dictionary<int, TSource> temp = new Dictionary<int, TSource>();
            int count = sheetMapper.ReadObjects(sheetName, temp);

            var filtered = temp.Where(n => function(n.Value));

            if (filtered.Any())
            {
                foreach (var current in filtered)
                {
                    buffer.Add(current);
                }
            }

            return count;
        }


        #region SheetWorker extension factories.
        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="sheetWorker"></param>
        public static IXLSheetWorker<TSource> AsXmlSheetWorker<TSource>(this IXLSheetWorker<TSource> sheetWorker)
            where TSource : class, new()
        {
            return new XLSheetMapper<TSource>(sheetWorker.HeaderRows, sheetWorker.PropertyMappers);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="sheetWorker"></param>
        /// <returns></returns>
        public static IXLSheetMapper<TSource> AsBiffSheetWorker<TSource>(this IXLSheetWorker<TSource> sheetWorker)
            where TSource : class, new()
        {
            return new XSheetMapper<TSource>(sheetWorker.HeaderRows, sheetWorker.PropertyMappers);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="sheetWorker"></param>
        /// <returns></returns>
        public static IXLSheetReader<TSource> AsXmlReader<TSource>(this IXLSheetWorker<TSource> sheetWorker)
            where TSource : class, new()
        {
            return new XLSheetReader<TSource>(sheetWorker.HeaderRows, sheetWorker.PropertyMappers);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="sheetWorker"></param>
        /// <returns></returns>
        public static IXLSheetReader<TSource> AsBiffReader<TSource>(this IXLSheetWorker<TSource> sheetWorker)
            where TSource : class, new()
        {
            return new XSheetReader<TSource>(sheetWorker.HeaderRows, sheetWorker.PropertyMappers);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="sheetWorker"></param>
        /// <returns></returns>
        public static IXLSheetWriter<TSource> AsXmlWriter<TSource>(this IXLSheetWorker<TSource> sheetWorker)
            where TSource : class, new()
        {
            return new XLSheetWriter<TSource>(sheetWorker.HeaderRows, sheetWorker.PropertyMappers);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="sheetWorker"></param>
        /// <returns></returns>
        public static IXLSheetWriter<TSource> AsBiffWriter<TSource>(this IXLSheetWorker<TSource> sheetWorker)
            where TSource : class, new()
        {
            return new XSheetWriter<TSource>(sheetWorker.HeaderRows, sheetWorker.PropertyMappers);
        }
        #endregion

    }
}

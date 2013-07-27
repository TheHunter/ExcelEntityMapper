using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelEntityMapper.Exceptions;

namespace ExcelEntityMapper.Impl
{
    /// <summary>
    /// /
    /// </summary>
    /// <typeparam name="TSource"></typeparam>
    public abstract class SheetWriter<TSource>
        : SheetBase, IXLSheetWriter<TSource>
        where TSource : class
    {
        private readonly List<IXLPropertyMapper<TSource>> propertyMappers = new List<IXLPropertyMapper<TSource>>();

        /// <summary>
        /// 
        /// </summary>
        /// <param name="headerRows"></param>
        /// <param name="zeroBase"></param>
        /// <param name="propertyMappers"></param>
        internal protected SheetWriter(int headerRows, bool zeroBase,
                                       IEnumerable<IXLPropertyMapper<TSource>> propertyMappers)
            : base(headerRows, zeroBase)
        {
            if (propertyMappers == null)
                throw new SheetParameterException("The sheet mapper cannot be null.", "propertyMappers");

            if (propertyMappers.Count(n => n.CustomType == MapperType.Key) == 0)
                throw new SheetParameterException("The SheetMapper must have at least a key PropertyMapper.", "propertyMappers");

            if (propertyMappers.Any( n=> n.OperationEnabled == SourceOperation.Write))
                throw new SheetParameterException("The current sheet writer must have no writeable property mappers.", "propertyMappers");

            var group = propertyMappers
                .GroupBy(n => n.ColumnIndex)
                .Where(n => n.Count() > 1)
                .Select(n => new KeyValuePair<int, int>(n.Key, n.Count()))
                ;

            if (group.Any())
                throw new SheetParameterException("The property mappers must have a unique column index, verify ColumnIndex property.", "propertyMappers");

            this.PropertyMappers = propertyMappers;
            this.LastColumn = this.PropertyMappers.Select(n => n.ColumnIndex).Max();
            this.FirstColumn = this.PropertyMappers.Select(n => n.ColumnIndex).Min();
        }

        /// <summary>
        /// 
        /// </summary>
        public Action<TSource> BeforeWriting { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public Action<TSource> AfterWriting { get; set; }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="instances"></param>
        /// <returns></returns>
        public int WriteObjects(IEnumerable<TSource> instances)
        {
            return this.WriteObjects(this.SheetName, instances);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="instances"></param>
        /// <returns></returns>
        public abstract int WriteObjects(string sheetName, IEnumerable<TSource> instances);

        /// <summary>
        /// 
        /// </summary>
        public IEnumerable<IXLPropertyMapper<TSource>> PropertyMappers
        {
            get { return propertyMappers; }
            private set
            {
                if (value == null || !value.Any())
                    throw new SheetParameterException("Column parameters cannot be null or empty.", "value");

                this.propertyMappers.AddRange(value);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="workbook"></param>
        public abstract void InjectWorkBook(IXLWorkBook workbook);
    }
}

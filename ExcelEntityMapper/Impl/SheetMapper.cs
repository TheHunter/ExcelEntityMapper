using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelEntityMapper.Exceptions;

namespace ExcelEntityMapper.Impl
{
    /// <summary>
    /// 
    /// </summary>
    /// <typeparam name="TSource"></typeparam>
    public abstract class SheetMapper<TSource>
        : SheetBase, IXLSheetMapper<TSource>
        where TSource : class, new()
    {
        private Action<TSource> beforeReading;
        private Action<TSource> afterReading;
        private Action<TSource> beforeWriting;
        private Action<TSource> afterWriting;
        private readonly List<IXLPropertyMapper<TSource>> propertyMappers = new List<IXLPropertyMapper<TSource>>();

        /// <summary>
        /// 
        /// </summary>
        /// <param name="headerRows"></param>
        /// <param name="zeroBase"></param>
        /// <param name="propertyMappers"></param>
        protected SheetMapper(int headerRows, bool zeroBase, IEnumerable<IXLPropertyMapper<TSource>> propertyMappers)
            : base(headerRows, zeroBase)
        {
            if (!propertyMappers.Any(n => n.CustomType == MapperType.Key))
                throw new SheetParameterException("The SheetMapper must have at least a key PropertyMapper", "propertyMappers");

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
        public Action<TSource> BeforeReading
        {
            get { return this.beforeReading; }
            set { this.beforeReading = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public Action<TSource> AfterReading
        {
            get { return this.afterReading; }
            set { this.afterReading = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public Action<TSource> BeforeWriting
        {
            get { return this.beforeWriting; }
            set { this.beforeWriting = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public Action<TSource> AfterWriting
        {
            get { return this.afterWriting; }
            set { this.afterWriting = value; }
        }

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
        /// <param name="buffer"></param>
        /// <returns></returns>
        public int ReadObjects(IDictionary<int, TSource> buffer)
        {
            return this.ReadObjects(this.SheetName, buffer);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="buffer"></param>
        /// <returns></returns>
        public abstract int ReadObjects(string sheetName, IDictionary<int, TSource> buffer);

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
    }
}

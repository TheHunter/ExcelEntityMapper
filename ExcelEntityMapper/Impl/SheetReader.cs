﻿using System;
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
    public abstract class SheetReader<TSource>
        : SheetBase, IXLSheetReader<TSource>
        where TSource : class, new()
    {
        private readonly List<IXLPropertyMapper<TSource>> propertyMappers = new List<IXLPropertyMapper<TSource>>();

        /// <summary>
        /// 
        /// </summary>
        /// <param name="headerRows"></param>
        /// <param name="zeroBase"></param>
        /// <param name="propertyMappers"></param>
        internal protected SheetReader(int headerRows, bool zeroBase,
                                        IEnumerable<IXLPropertyMapper<TSource>> propertyMappers)
            : base(headerRows, zeroBase)
        {
            if (propertyMappers == null)
                throw new SheetParameterException("The sheet mapper cannot be null.", "propertyMappers");

            if (propertyMappers.Count(n => n.CustomType == MapperType.Key) == 0)
                throw new SheetParameterException("The SheetMapper must have at least a key PropertyMapper.", "propertyMappers");

            if (propertyMappers.Any(n => n.OperationEnabled == SourceOperation.Read))
                throw new SheetParameterException("The current sheet reader must have no readable property mappers.", "propertyMappers");

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
        public Action<TSource> BeforeReading { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public Action<TSource> AfterReading { get; set; }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public abstract int GetIndexFirstRow(string sheetName);
        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <returns></returns>
        public TSource ReadObject(int rowIndex)
        {
            return this.ReadObject(this.SheetName, rowIndex);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="rowIndex"></param>
        /// <returns></returns>
        public abstract TSource ReadObject(string sheetName, int rowIndex);

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
        /// <param name="workbook"></param>
        public abstract void InjectWorkBook(IXLWorkBook workbook);
    }
}

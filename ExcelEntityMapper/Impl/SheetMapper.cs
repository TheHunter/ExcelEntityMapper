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
        private readonly List<IXLPropertyMapper<TSource>> propertyMappers = new List<IXLPropertyMapper<TSource>>();

        ///// <summary>
        ///// 
        ///// </summary>
        ///// <param name="indexkeyColumn"></param>
        ///// <param name="parameters"></param>
        //public SheetMapper(int indexkeyColumn, IEnumerable<IXLPropertyMapper<TSource>> propertyMappers)
        //    : this(indexkeyColumn, false, propertyMappers)
        //{
        //}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="indexkeyColumn"></param>
        /// <param name="hasHeader"></param>
        /// <param name="zeroBase"></param>
        /// <param name="propertyMappers"></param>
        public SheetMapper(int indexkeyColumn, bool hasHeader, bool zeroBase, IEnumerable<IXLPropertyMapper<TSource>> propertyMappers)
            : base(indexkeyColumn, hasHeader, zeroBase)
        {
            this.PropertyMappers = propertyMappers;
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

    }
}

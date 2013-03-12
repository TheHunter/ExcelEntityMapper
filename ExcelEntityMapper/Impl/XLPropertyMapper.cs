using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Linq.Expressions;

namespace ExcelEntityMapper.Impl
{
    /// <summary>
    /// 
    /// </summary>
    /// <typeparam name="TSource"></typeparam>
    public class XLPropertyMapper<TSource>
        : IXLPropertyMapper<TSource>
        where TSource : class
    {
        private int _ColumnIndex = -1;
        private Func<TSource, string> _ToExcelFormat = null;
        private Action<TSource, string> _ToPropertyFormat = null;
        private string _ColumnHeader = null;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="column"></param>
        /// <param name="toPropertyFormat"></param>
        /// <param name="toExcelFormat"></param>
        public XLPropertyMapper(int column, Action<TSource, string> toPropertyFormat, Expression<Func<TSource, string>> toExcelFormat)
		{
            this.ColumnIndex = column;
            this.ToPropertyFormat = toPropertyFormat;
            this.ToExcelFormat = toExcelFormat.Compile();

            try
            {
                var temp = SourceHelper.GetMemberInfo(toExcelFormat);
                this.ColumnHeader = temp.Key;
            }
            catch (Exception)
            {
                // non si fa nulla se non si recupera il nome completo della property da mappare.
                this.ColumnHeader = "No Property header";
            }
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="column"></param>
        /// <param name="columnHeader"></param>
        /// <param name="toPropertyFormat"></param>
        /// <param name="toExcelFormat"></param>
        public XLPropertyMapper(int column, string columnHeader, Action<TSource, string> toPropertyFormat, Expression<Func<TSource, string>> toExcelFormat)
        {
            this.ColumnIndex = column;
            this.ToPropertyFormat = toPropertyFormat;
            this.ToExcelFormat = toExcelFormat.Compile();
            this.ColumnHeader = columnHeader;
        }

        /// <summary>
        /// 
        /// </summary>
        public int ColumnIndex
        {
            protected set
            {
                if (value < 1)
                    throw new IndexOutOfRangeException("Column index cannot be less than one.");
                else
                    this._ColumnIndex = value;
            }
            get { return this._ColumnIndex; }
        }
        
        /// <summary>
        /// 
        /// </summary>
        public string ColumnHeader
        {
            get { return this._ColumnHeader; }
            private set
            {
                if (string.IsNullOrEmpty(value))
                {
                    this._ColumnHeader = "[No property header]";
                }
                else
                {
                    this._ColumnHeader = value.Trim();
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public Action<TSource, string> ToPropertyFormat
        {
            get { return this._ToPropertyFormat; }
            private set { this._ToPropertyFormat = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public Func<TSource, string> ToExcelFormat
        {
            get { return this._ToExcelFormat; }
            private set { this._ToExcelFormat = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            if (obj != null && obj is XLPropertyMapper<TSource>)
            {
                return this.GetHashCode() == obj.GetHashCode();
            }
            return false;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return ((13 * typeof(TSource).Name.GetHashCode()) - this._ColumnIndex);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return string.Format("Property of <{0}>, Column index: {1}, Column name: {2}", typeof(TSource).Name, this.ColumnIndex, this.ColumnHeader);
        }
    }
}

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
	public class PropertyMapper<TSource>
		: IXLPropertyMapper<TSource>
		where TSource : class
	{
		private int columnIndex = -1;
		private string columnHeader;
	    private MapperType propMapper = MapperType.Simple;
		private Func<TSource, string> toExcelFormat;
		private Action<TSource, string> toPropertyFormat;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="column"></param>
        /// <param name="toPropertyFormat"></param>
        /// <param name="toExcelFormat"></param>
        public PropertyMapper(int column, Action<TSource, string> toPropertyFormat,
                              Expression<Func<TSource, string>> toExcelFormat)
            : this(
                column, MapperType.Simple, SourceHelper.GetDefaultMemberName(toExcelFormat, "NoPropertyHeader"),
                toPropertyFormat, toExcelFormat)
        {
            
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="column"></param>
		/// <param name="mapperType"></param>
		/// <param name="toPropertyFormat"></param>
		/// <param name="toExcelFormat"></param>
		public PropertyMapper(int column, MapperType mapperType, Action<TSource, string> toPropertyFormat,
		                      Expression<Func<TSource, string>> toExcelFormat)
		    : this(
		        column, mapperType, SourceHelper.GetDefaultMemberName(toExcelFormat, "NoPropertyHeader"), toPropertyFormat,
		        toExcelFormat)
		{
            //this.ColumnIndex = column;
            //this.ToPropertyFormat = toPropertyFormat;
            //this.ToExcelFormat = toExcelFormat.Compile();

            //try
            //{
            //    var temp = SourceHelper.GetMemberInfo(toExcelFormat);
            //    this.ColumnHeader = temp.Key;
            //}
            //catch (Exception)
            //{
            //    // non si fa nulla se non si recupera il nome completo della property da mappare.
            //    this.ColumnHeader = "No Property header";
            //}
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="column"></param>
        /// <param name="columnHeader"></param>
        /// <param name="toPropertyFormat"></param>
        /// <param name="toExcelFormat"></param>
        public PropertyMapper(int column, string columnHeader,
                              Action<TSource, string> toPropertyFormat, Expression<Func<TSource, string>> toExcelFormat)
            :this(column, MapperType.Simple, columnHeader, toPropertyFormat, toExcelFormat)
        {
            
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="column"></param>
		/// <param name="mapperType"></param>
		/// <param name="columnHeader"></param>
		/// <param name="toPropertyFormat"></param>
		/// <param name="toExcelFormat"></param>
		public PropertyMapper(int column, MapperType mapperType, string columnHeader, Action<TSource, string> toPropertyFormat,
		                      Expression<Func<TSource, string>> toExcelFormat)
		{
            this.ColumnIndex = column;
            this.propMapper = mapperType;
            this.ColumnHeader = columnHeader;
            this.ToPropertyFormat = toPropertyFormat;
            this.ToExcelFormat = toExcelFormat.Compile();
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
				this.columnIndex = value;
			}
			get { return this.columnIndex; }
		}
		
		/// <summary>
		/// 
		/// </summary>
		public string ColumnHeader
		{
			get { return this.columnHeader; }
			private set
			{
				if (string.IsNullOrEmpty(value))
				{
					this.columnHeader = "[No property header]";
				}
				else
				{
					this.columnHeader = value.Trim();
				}
			}
		}

        /// <summary>
        /// 
        /// </summary>
	    public MapperType CustomType
	    {
            get { return this.propMapper; }
	    }

		/// <summary>
		/// 
		/// </summary>
		public Action<TSource, string> ToPropertyFormat
		{
			get { return this.toPropertyFormat; }
			private set { this.toPropertyFormat = value; }
		}

		/// <summary>
		/// 
		/// </summary>
		public Func<TSource, string> ToExcelFormat
		{
			get { return this.toExcelFormat; }
			private set { this.toExcelFormat = value; }
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="obj"></param>
		/// <returns></returns>
		public override bool Equals(object obj)
		{
			if (obj != null && obj is PropertyMapper<TSource>)
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
			return ((13 * typeof(TSource).Name.GetHashCode()) - this.columnIndex);
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

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Linq.Expressions;
using ExcelEntityMapper.Exceptions;

namespace ExcelEntityMapper.Impl
{
	/// <summary>
	/// 
	/// </summary>
	/// <typeparam name="TSource"></typeparam>
	public class PropertyMapper<TSource>
		: ColumnProperty, IXLPropertyMapper<TSource>
		where TSource : class
	{
		private readonly Action<TSource, string> toPropertyFormat;
		private readonly Func<TSource, object> toExcelFormat;

		
		/// <summary>
		/// 
		/// </summary>
		/// <param name="column"></param>
		/// <param name="toPropertyFormat"></param>
		/// <param name="toExcelFormat"></param>
		public PropertyMapper(int column, Action<TSource, string> toPropertyFormat,
							  Expression<Func<TSource, object>> toExcelFormat)
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
							  Expression<Func<TSource, object>> toExcelFormat)
			: this(
				column, mapperType, SourceHelper.GetDefaultMemberName(toExcelFormat, "NoPropertyHeader"), toPropertyFormat,
				toExcelFormat)
		{
		}
		
		/// <summary>
		/// 
		/// </summary>
		/// <param name="column"></param>
		/// <param name="columnHeader"></param>
		/// <param name="toPropertyFormat"></param>
		/// <param name="toExcelFormat"></param>
		public PropertyMapper(int column, string columnHeader,
							  Action<TSource, string> toPropertyFormat, Expression<Func<TSource, object>> toExcelFormat)
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
	    /// <exception cref="WrongParameterException"></exception>
	    public PropertyMapper(int column, MapperType mapperType, string columnHeader, Action<TSource, string> toPropertyFormat,
							  Expression<Func<TSource, object>> toExcelFormat)
			: this(column, mapperType, columnHeader, SourceOperation.ReadWrite)
		{
			
			if (toPropertyFormat == null)
				throw new WrongParameterException("The action for setting property source value cannot be null.", "toPropertyFormat");

			if (toExcelFormat == null)
				throw new WrongParameterException("The expression for getting property source value cannot be null.", "toExcelFormat");

			this.toPropertyFormat = toPropertyFormat;
			this.toExcelFormat = toExcelFormat.Compile();
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="column"></param>
		/// <param name="mapperType"></param>
		/// <param name="columnHeader"></param>
		/// <param name="operationEnabled"></param>
		private PropertyMapper(int column, MapperType mapperType, string columnHeader, SourceOperation operationEnabled)
			: base(column, mapperType, columnHeader, operationEnabled)
		{
		}

		#region

	    /// <summary>
	    /// 
	    /// </summary>
	    /// <param name="column"></param>
	    /// <param name="mapperType"></param>
	    /// <param name="columnHeader"></param>
	    /// <param name="toExcelFormat"></param>
	    /// <exception cref="WrongParameterException"></exception>
	    internal PropertyMapper(int column, MapperType mapperType, string columnHeader, Expression<Func<TSource, object>> toExcelFormat)
			: base(column, mapperType, columnHeader, SourceOperation.Read)
		{
			if (toExcelFormat == null)
				throw new WrongParameterException("The expression for getting property source value cannot be null.", "toExcelFormat");

			this.toExcelFormat = toExcelFormat.Compile();
		}

	    /// <summary>
	    /// 
	    /// </summary>
	    /// <param name="column"></param>
	    /// <param name="mapperType"></param>
	    /// <param name="columnHeader"></param>
	    /// <param name="toPropertyFormat"></param>
	    /// <exception cref="WrongParameterException"></exception>
	    internal PropertyMapper(int column, MapperType mapperType, string columnHeader, Action<TSource, string> toPropertyFormat)
			: base(column, mapperType, columnHeader, SourceOperation.Write)
		{
			if (toPropertyFormat == null)
				throw new WrongParameterException("The action for setting property source value cannot be null.", "toPropertyFormat");

			this.toPropertyFormat = toPropertyFormat;
		}

		#endregion

		/// <summary>
		/// 
		/// </summary>
		public Action<TSource, string> ToPropertyFormat
		{
			get { return this.toPropertyFormat; }
		}

		/// <summary>
		/// 
		/// </summary>
		public Func<TSource, object> ToExcelFormat
		{
			get { return this.toExcelFormat; }
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="obj"></param>
		/// <returns></returns>
		public override bool Equals(object obj)
		{
			if (obj == null)
				return false;

			if (obj is PropertyMapper<TSource>)
				return this.GetHashCode() == obj.GetHashCode();
			
			return false;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <returns></returns>
		public override int GetHashCode()
		{
			return ((13 * typeof(TSource).Name.GetHashCode()) - this.ColumnIndex);
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

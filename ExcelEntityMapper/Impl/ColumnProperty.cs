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
    public class ColumnProperty
        : IXLProperty
    {
        private readonly int columnIndex = -1;
        private readonly string columnHeader;
        private readonly SourceOperation operationEnabled;
        private readonly MapperType propMapper = MapperType.Simple;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="column"></param>
        /// <param name="mapperType"></param>
        /// <param name="columnHeader"></param>
        /// <param name="operationEnabled"></param>
        public ColumnProperty(int column, MapperType mapperType, string columnHeader, SourceOperation operationEnabled)
        {
            if (column < 1)
                throw new WrongParameterException("Column index cannot be less than one.", "columnIndex");

            this.columnIndex = column;
            this.propMapper = mapperType;
            this.columnHeader = columnHeader == null || columnHeader.Trim().Equals(string.Empty) ? "[No property header]" : columnHeader.Trim();
            this.operationEnabled = operationEnabled;
        }

        /// <summary>
        /// 
        /// </summary>
        public int ColumnIndex
        {
            get { return columnIndex; }
        }

        /// <summary>
        /// 
        /// </summary>
        public string ColumnHeader
        {
            get { return columnHeader; }
        }

        /// <summary>
        /// 
        /// </summary>
        public MapperType CustomType
        {
            get { return propMapper; }
        }

        /// <summary>
        /// 
        /// </summary>
        public SourceOperation OperationEnabled
        {
            get { return operationEnabled; }
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

            if (obj is ColumnProperty)
                return this.GetHashCode() == obj.GetHashCode();

            return false;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return (-11 * ColumnIndex);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return string.Format("Header: {0}, Index: {1}", this.columnHeader, this.ColumnIndex);
        }
    }
}

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
    public abstract class SheetBase
        : IXLSheet
    {
        private readonly int indexKeyColumn;
        private readonly bool hasHeader = true;
        private readonly int offset;
        private string sheetName;
        private int lastColumn = 1;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="indexkeyColumn"></param>
        protected SheetBase(int indexkeyColumn)
            :this(indexkeyColumn, false, false)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="indexkeyColumn"></param>
        /// <param name="zeroBase"></param>
        protected SheetBase(int indexkeyColumn, bool zeroBase)
            : this(indexkeyColumn, false, zeroBase)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="indexkeyColumn"></param>
        /// <param name="hasHeader"></param>
        /// <param name="zeroBase"></param>
        protected SheetBase(int indexkeyColumn, bool hasHeader, bool zeroBase)
        {
            if (indexkeyColumn < 1)
                throw new SheetParameterException(
                    string.Format(
                        "L'indice della colonna in cui si trova la chiave non puo' essere inferiore a uno, valore: {0}",
                        indexkeyColumn), "indexkeyColumn");

            this.indexKeyColumn = indexkeyColumn;
            this.hasHeader = hasHeader;

            if (zeroBase) this.offset = 1;
        }

        /// <summary>
        /// 
        /// </summary>
        public string SheetName
        {
            get { return this.sheetName; }
            set
            {
                if (string.IsNullOrEmpty(value))
                    throw new SheetParameterException("The sheet name of the calling object cannot be null or empty.", "value");
                
                this.sheetName = value.Trim();
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public int IndexKeyColumn
        {
            get { return this.indexKeyColumn; }
        }

        /// <summary>
        /// 
        /// </summary>
        public int LastColumn
        {
            get { return this.lastColumn; }
            protected set
            {
                if (value < 1)
                    throw new SheetParameterException("Index of last column cannot be less than one.", "value");
                this.lastColumn = value;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public bool HasHeader
        {
            get { return this.hasHeader; }
        }

        /// <summary>
        /// 
        /// </summary>
        public bool ZeroBase
        {
            get { return offset == 1; }
        }

        /// <summary>
        /// 
        /// </summary>
        protected int Offset
        {
            get { return this.offset; }
        }

    }
}

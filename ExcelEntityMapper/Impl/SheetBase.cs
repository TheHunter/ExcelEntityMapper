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
        private readonly int headerRows;
        private readonly int offset;
        private string sheetName;
        private int lastColumn = 1;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="indexkeyColumn"></param>
        protected SheetBase(int indexkeyColumn)
            :this(indexkeyColumn, 0, false)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="indexkeyColumn"></param>
        /// <param name="zeroBase"></param>
        protected SheetBase(int indexkeyColumn, bool zeroBase)
            : this(indexkeyColumn, 0, zeroBase)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="indexkeyColumn"></param>
        /// <param name="headerRows"></param>
        /// <param name="zeroBase"></param>
        protected SheetBase(int indexkeyColumn, int headerRows, bool zeroBase)
        {
            if (indexkeyColumn < 1)
                throw new SheetParameterException(
                    string.Format(
                        "L'indice della colonna in cui si trova la chiave non puo' essere inferiore a uno, valore: {0}",
                        indexkeyColumn), "indexkeyColumn");

            if (headerRows < 0)
                headerRows = 0;

            this.indexKeyColumn = indexkeyColumn;
            this.headerRows = headerRows;

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
            get { return headerRows > 0; }
        }

        /// <summary>
        /// 
        /// </summary>
        public int HeaderRows
        {
            get { return this.headerRows; }
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

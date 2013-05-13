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
        private readonly int headerRows;
        private readonly int offset;
        private string sheetName;
        private int lastColumn = 1;
        private int firstColumn = 1;

        /// <summary>
        /// 
        /// </summary>
        protected SheetBase()
            :this(0, false)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="zeroBase"></param>
        protected SheetBase(bool zeroBase)
            : this(0, zeroBase)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="headerRows"></param>
        /// <param name="zeroBase"></param>
        protected SheetBase(int headerRows, bool zeroBase)
        {
            if (headerRows < 0)
                headerRows = 0;

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
        public int FirstColumn
        {
            get { return this.FirstColumn; }
            protected set
            {
                if (value < 1)
                    throw new SheetParameterException("Index of first column cannot be less than one.", "value");
                this.firstColumn = value;
            }
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

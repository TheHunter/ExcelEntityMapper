using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelEntityMapper.Impl
{
    /// <summary>
    /// 
    /// </summary>
    public abstract class XLSheet
        : IXLSheet
    {
        private readonly int _IndexKeyColumn;
        private readonly bool _HasHeader = true;
        private string _SheetName = null;
        private int _LastColumn = 1;
        private int _Offset = 0;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="indexkeyColumn"></param>
        protected XLSheet(int indexkeyColumn)
            :this(indexkeyColumn, false, false)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="indexkeyColumn"></param>
        /// <param name="zeroBase"></param>
        protected XLSheet(int indexkeyColumn, bool zeroBase)
            : this(indexkeyColumn, false, zeroBase)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="indexkeyColumn"></param>
        /// <param name="hasHeader"></param>
        /// <param name="zeroBase"></param>
        protected XLSheet(int indexkeyColumn, bool hasHeader, bool zeroBase)
        {
            if (indexkeyColumn < 1)
                throw new ArgumentException(string.Format("L'indice della colonna in cui si trova la chiave non puo' essere inferiore a uno, valore: {0}", indexkeyColumn));
            this._IndexKeyColumn = indexkeyColumn;
            this._HasHeader = hasHeader;

            if (zeroBase) this._Offset = 1;
        }

        /// <summary>
        /// 
        /// </summary>
        public string SheetName
        {
            get { return this._SheetName; }
            set
            {
                if (string.IsNullOrEmpty(value))
                    throw new ArgumentNullException("The sheet name of the calling object cannot be null or empty.");
                
                this._SheetName = value.Trim();
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public int IndexKeyColumn
        {
            get { return this._IndexKeyColumn; }
        }

        /// <summary>
        /// 
        /// </summary>
        public int LastColumn
        {
            get { return this._LastColumn; }
            protected set
            {
                if (value < 1)
                    throw new ArgumentException("Index of last column cannot be less than one.");
                this._LastColumn = value;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public bool HasHeader
        {
            get { return this._HasHeader; }
        }

        /// <summary>
        /// 
        /// </summary>
        public bool ZeroBase
        {
            get { return _Offset == 1; }
        }

        /// <summary>
        /// 
        /// </summary>
        protected int Offset
        {
            get { return this._Offset; }
        }

    }
}

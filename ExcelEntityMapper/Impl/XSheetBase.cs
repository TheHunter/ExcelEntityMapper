using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelEntityMapper.Impl
{
    /// <summary>
    /// 
    /// </summary>
    public abstract class XSheetBase
        : XLSheet
    {
         /// <summary>
        /// 
        /// </summary>
        /// <param name="indexkeyColumn"></param>
        protected XSheetBase(int indexkeyColumn)
            :this(indexkeyColumn, false)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="indexkeyColumn"></param>
        /// <param name="hasHeader"></param>
        protected XSheetBase(int indexkeyColumn, bool hasHeader)
            :this(indexkeyColumn, hasHeader, false)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="indexkeyColumn"></param>
        /// <param name="hasHeader"></param>
        /// <param name="zeroBase"></param>
        protected XSheetBase(int indexkeyColumn, bool hasHeader, bool zeroBase)
            :base(indexkeyColumn, hasHeader, zeroBase)
        {

        }
    }
}

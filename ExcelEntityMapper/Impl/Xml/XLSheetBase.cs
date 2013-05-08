using ClosedXML.Excel;

namespace ExcelEntityMapper.Impl.Xml
{
    /// <summary>
    /// 
    /// </summary>
    internal abstract class XLSheetBase
        : XLSheet
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="indexkeyColumn"></param>
        protected XLSheetBase(int indexkeyColumn)
            :this(indexkeyColumn, false)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="indexkeyColumn"></param>
        /// <param name="hasHeader"></param>
        protected XLSheetBase(int indexkeyColumn, bool hasHeader)
            :base(indexkeyColumn, hasHeader)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="index"></param>
        /// <param name="color"></param>
        protected abstract void ChangeBackGround(int index, XLColor color);
    }
}

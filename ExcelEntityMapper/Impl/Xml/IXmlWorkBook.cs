using ClosedXML.Excel;

namespace ExcelEntityMapper.Impl.Xml
{
    /// <summary>
    /// 
    /// </summary>
    internal interface IXmlWorkBook
        : IXLWorkBook
    {
        /// <summary>
        /// 
        /// </summary>
        XLWorkbook WorkBook { get; }
    }
}

using NPOI.HSSF.UserModel;

namespace ExcelEntityMapper.Impl.BIFF
{
    /// <summary>
    /// 
    /// </summary>
    internal interface IBiffWorkBook
        : IXLWorkBook
    {
        HSSFWorkbook WorkBook { get; }
    }
}

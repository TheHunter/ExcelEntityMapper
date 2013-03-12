using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using ClosedXML.Excel;

namespace ExcelEntityMapper.Impl
{
    /// <summary>
    /// 
    /// </summary>
    internal class XLWorkBook
        //: IXLWorkBook
    {
        private XLWorkbook _Workbook = null;

        /// <summary>
        /// 
        /// </summary>
        public XLWorkBook()
        {
            _Workbook = new XLWorkbook();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="inputStream"></param>
        public XLWorkBook(Stream inputStream)
        {
            _Workbook = new XLWorkbook(inputStream);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public bool ExistsWorkSheet(string sheetName)
        {
            return this._Workbook.Worksheets.Count(n => n.Name == sheetName) > 0;
        }
        
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public IEnumerable<string> GetSheetNames()
        {
            return this._Workbook.Worksheets.Select(n => n.Name);
        }

        /// <summary>
        /// 
        /// </summary>
        internal XLWorkbook Workbook
        {
            get { return this._Workbook; }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public bool AddSheet(string sheetName)
        {
            return this._Workbook.AddWorksheet(sheetName) != null;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public bool RemoveSheet(string sheetName)
        {
            this._Workbook.Worksheets.Delete(sheetName);
            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        public Stream Save()
        {
            MemoryStream stream = new MemoryStream();
            this._Workbook.SaveAs(stream);
            return stream;
        }
    }
}

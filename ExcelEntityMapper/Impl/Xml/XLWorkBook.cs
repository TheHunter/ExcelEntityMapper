using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using ClosedXML.Excel;
using ExcelEntityMapper.Exceptions;

namespace ExcelEntityMapper.Impl.Xml
{
    /// <summary>
    /// 
    /// </summary>
    public class XLWorkBook
        : IXLWorkBook
    {
        private readonly XLWorkbook workbook;

        /// <summary>
        /// 
        /// </summary>
        public XLWorkBook()
        {
            workbook = new XLWorkbook();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="inputStream"></param>
        public XLWorkBook(byte[] inputStream)
        {
            try
            {
                workbook = new XLWorkbook(new MemoryStream(inputStream)); ;
            }
            catch (Exception ex)
            {
                throw new WorkBookException("Error on reading the stream input.", ex);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="inputStream"></param>
        public XLWorkBook(Stream inputStream)
        {
            try
            {
                workbook = new XLWorkbook(inputStream);
            }
            catch (Exception ex)
            {
                throw new WorkBookException("Error on reading the stream input.", ex);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public bool ExistsWorkSheet(string sheetName)
        {
            return this.workbook.Worksheets.Count(n => n.Name == sheetName) > 0;
        }
        
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public IEnumerable<string> GetSheetNames()
        {
            return this.workbook.Worksheets.Select(n => n.Name);
        }

        /// <summary>
        /// 
        /// </summary>
        internal XLWorkbook Workbook
        {
            get { return this.workbook; }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public bool AddSheet(string sheetName)
        {
            return this.workbook.AddWorksheet(sheetName) != null;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public bool RemoveSheet(string sheetName)
        {
            this.workbook.Worksheets.Delete(sheetName);
            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        public Stream Save()
        {
            MemoryStream stream = new MemoryStream();
            this.workbook.SaveAs(stream);
            return stream;
        }
    }
}

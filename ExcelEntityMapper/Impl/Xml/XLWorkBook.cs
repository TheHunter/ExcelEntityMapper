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
        private readonly XLWorkbook workBook;

        /// <summary>
        /// 
        /// </summary>
        public XLWorkBook()
        {
            workBook = new XLWorkbook();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="inputStream"></param>
        public XLWorkBook(byte[] inputStream)
        {
            try
            {
                workBook = new XLWorkbook(new MemoryStream(inputStream));
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
                workBook = new XLWorkbook(inputStream);
            }
            catch (Exception ex)
            {
                throw new WorkBookException("Error on reading the stream input.", ex);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="workBook"></param>
        public XLWorkBook(XLWorkbook workBook)
        {
            this.workBook = workBook;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public bool ExistsWorkSheet(string sheetName)
        {
            return this.workBook.Worksheets.Count(n => n.Name == sheetName) > 0;
        }
        
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public IEnumerable<string> GetSheetNames()
        {
            return this.workBook.Worksheets.Select(n => n.Name);
        }

        /// <summary>
        /// 
        /// </summary>
        internal XLWorkbook WorkBook
        {
            get { return this.workBook; }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public bool AddSheet(string sheetName)
        {
            if (sheetName == null || sheetName.Trim().Equals(string.Empty))
                return false;

            if (!this.ExistsWorkSheet(sheetName))
            {
                this.workBook.AddWorksheet(sheetName);
                return true;
            }
            return false;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public bool RemoveSheet(string sheetName)
        {
            if (sheetName == null || sheetName.Trim().Equals(string.Empty))
                return false;

            this.workBook.Worksheets.Delete(sheetName);
            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        public ExcelFormat Format
        {
            get { return ExcelFormat.XML; }
        }

        /// <summary>
        /// 
        /// </summary>
        public Stream Save()
        {
            MemoryStream stream = new MemoryStream();
            this.workBook.SaveAs(stream);
            return stream;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="propertyMappers"></param>
        /// <returns></returns>
        public IXLSheetWorker<TSource> MakeSheetMapper<TSource>(IEnumerable<IXLPropertyMapper<TSource>> propertyMappers)
            where TSource : class, new()
        {
            var mapper = new XLSheetMapper<TSource>(propertyMappers);
            IXLWorkBookProvider<TSource> a = mapper;
            a.WorkBook = this.WorkBook;
            return mapper;
        }
    }
}

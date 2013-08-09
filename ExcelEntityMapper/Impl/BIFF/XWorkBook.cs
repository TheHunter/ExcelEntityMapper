using System;
using System.Collections.Generic;
using ExcelEntityMapper.Exceptions;
using System.IO;
using System.Linq;
using NPOI.HSSF.UserModel;

namespace ExcelEntityMapper.Impl.BIFF
{
    /// <summary>
    /// An object class which manage a standard WorkBook of different formats.
    /// </summary>
    public class XWorkBook
        : IXLWorkBook
    {
        private readonly HSSFWorkbook workBook;

        /// <summary>
        /// Instance a empty XWorkBook.
        /// </summary>
        public XWorkBook()
        {
            workBook = new HSSFWorkbook();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="inputStream"></param>
        public XWorkBook(byte[] inputStream)
        {
            try
            {
                workBook = new HSSFWorkbook(new MemoryStream(inputStream));
            }
            catch (Exception ex)
            {
                throw new WorkBookException("Error on reading the stream input.", ex);
            }
        }

        /// <summary>
        /// Intance a new WorkBook from the stream input
        /// </summary>
        /// <param name="inputStream">a stream which contains the file to read.</param>
        public XWorkBook(Stream inputStream)
        {
            try
            {
                workBook = new HSSFWorkbook(inputStream);
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
        internal XWorkBook(HSSFWorkbook workBook)
        {
            this.workBook = workBook;
        }

        /// <summary>
        /// Verify if It exists a worksheet with the given name.
        /// </summary>
        /// <param name="sheetName">the worksheet sheetname to find.</param>
        /// <returns>return true if it exists the worksheet.</returns>
        public bool ExistsWorkSheet(string sheetName)
        {
            return this.workBook.GetSheetIndex(sheetName) != -1;
        }

        /// <summary>
        /// get all sheet names present into the calling WorkBook.
        /// </summary>
        /// <returns>return all sheet names.</returns>
        public IEnumerable<string> GetSheetNames()
        {
            List<string> names = new List<string>();
            for (int i = 0; i < this.workBook.NumberOfSheets; i++)
            {
                names.Add(this.workBook.GetSheetName(i));
            }
            return names;
        }

        /// <summary>
        /// 
        /// </summary>
        internal HSSFWorkbook WorkBook
        {
            get { return this.workBook; }
        }

        /// <summary>
        /// Try to add a worksheet with the given name.
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public bool AddSheet(string sheetName)
        {
            if (sheetName == null || sheetName.Trim().Equals(string.Empty))
                return false;

            if (!this.ExistsWorkSheet(sheetName))
            {
                this.workBook.CreateSheet(sheetName);
                return true;
            }
            return false;
        }

        /// <summary>
        /// removes a worksheet which matches with the given sheet name.
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public bool RemoveSheet(string sheetName)
        {
            if (sheetName == null || sheetName.Trim().Equals(string.Empty))
                return false;

            if (this.ExistsWorkSheet(sheetName))
            {
                this.workBook.RemoveSheetAt(this.workBook.GetSheetIndex(sheetName));
                return true;
            }
            return false;
        }

        /// <summary>
        /// 
        /// </summary>
        public ExcelFormat Format
        {
            get { return ExcelFormat.BIFF; }
        }

        /// <summary>
        /// Saves the calling WorkBook into stream. 
        /// </summary>
        /// <returns>return a stream wich contains the calling WorkBook.</returns>
        public Stream Save()
        {
            MemoryStream stream = new MemoryStream();
            this.workBook.Write(stream);
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
            var mapper = new XSheetMapper<TSource>(propertyMappers);
            IXWorkBookProvider<TSource> a = mapper;
            a.WorkBook = this.WorkBook;
            return mapper;
        }
    }
}

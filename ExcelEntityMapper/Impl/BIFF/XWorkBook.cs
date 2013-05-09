using System;
using System.Collections.Generic;
using ExcelEntityMapper.Exceptions;
using System.IO;
using System.Linq;
using NPOI.HSSF.UserModel;

namespace ExcelEntityMapper.Impl.BIFF
{
    /// <summary>
    /// An object class which manage a standard workbook of different formats.
    /// </summary>
    public class XWorkBook
        : IXLWorkBook
    {
        private readonly HSSFWorkbook workbook;

        /// <summary>
        /// Instance a empty XWorkBook.
        /// </summary>
        public XWorkBook()
        {
            workbook = new HSSFWorkbook();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="inputStream"></param>
        public XWorkBook(byte[] inputStream)
        {
            try
            {
                workbook = new HSSFWorkbook(new MemoryStream(inputStream));
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
                workbook = new HSSFWorkbook(inputStream);
            }
            catch (Exception ex)
            {
                throw new WorkBookException("Error on reading the stream input.", ex);
            }
        }

        /// <summary>
        /// Verify if It exists a worksheet with the given name.
        /// </summary>
        /// <param name="sheetName">the worksheet sheetname to find.</param>
        /// <returns>return true if it exists the worksheet.</returns>
        public bool ExistsWorkSheet(string sheetName)
        {
            return this.workbook.GetSheetIndex(sheetName) != -1;
        }

        /// <summary>
        /// get all sheet names present into the calling workbook.
        /// </summary>
        /// <returns>return all sheet names.</returns>
        public IEnumerable<string> GetSheetNames()
        {
            List<string> names = new List<string>();
            for (int i = 0; i < this.workbook.NumberOfSheets; i++)
            {
                names.Add(this.workbook.GetSheetName(i));
            }
            return names;
        }

        /// <summary>
        /// 
        /// </summary>
        internal HSSFWorkbook Workbook
        {
            get { return this.workbook; }
        }

        /// <summary>
        /// Try to add a worksheet with the given name.
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public bool AddSheet(string sheetName)
        {
            if (!this.ExistsWorkSheet(sheetName))
            {
                this.workbook.CreateSheet(sheetName);
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
            if (this.ExistsWorkSheet(sheetName))
            {
                this.workbook.RemoveSheetAt(this.workbook.GetSheetIndex(sheetName));
                return true;
            }
            return false;
        }

        /// <summary>
        /// Saves the calling WorkBook into stream. 
        /// </summary>
        /// <returns>return a stream wich contains the calling workbook.</returns>
        public Stream Save()
        {
            MemoryStream stream = new MemoryStream();
            this.workbook.Write(stream);
            return stream;
        }
    }
}

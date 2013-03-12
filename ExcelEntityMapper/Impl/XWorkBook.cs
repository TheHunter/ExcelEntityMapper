using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SmartXLS;
using System.IO;

namespace ExcelEntityMapper.Impl
{
    /// <summary>
    /// An object class which manage a standard workbook of different formats.
    /// </summary>
    public class XWorkBook
        : IXLWorkBook
    {
        private WorkBook _Workbook = null;

        /// <summary>
        /// Instance a empty XWorkBook.
        /// </summary>
        public XWorkBook()
        {
            _Workbook = new WorkBook();
        }

        /// <summary>
        /// Intance a new WorkBook from the stream input
        /// </summary>
        /// <param name="inputStream">a stream which contains the file to read.</param>
        /// <param name="format">the excel format to read.</param>
        public XWorkBook(Stream inputStream, ExcelFormat format)
        {
            _Workbook = new WorkBook();
            switch (format)
            {
                case ExcelFormat.Xls:
                    {
                        this._Workbook.read(inputStream);
                        break;
                    }
                case ExcelFormat.Xlsx:
                    {
                        this._Workbook.readXLSX(inputStream);
                        break;
                    }
                default:
                    {
                        throw new NotImplementedException("Workbook format not implemented.");
                    }
            }
        }

        /// <summary>
        /// Verify if It exists a worksheet with the given name.
        /// </summary>
        /// <param name="sheetName">the worksheet sheetname to find.</param>
        /// <returns>return true if it exists the worksheet.</returns>
        public bool ExistsWorkSheet(string sheetName)
        {
            return _Workbook.findSheetByName(sheetName) != -1;
        }

        /// <summary>
        /// get all sheet names present into the calling workbook.
        /// </summary>
        /// <returns>return all sheet names.</returns>
        public IEnumerable<string> GetSheetNames()
        {
            List<string> sheetNames = new List<string>();
            for (int index = 0; index < this._Workbook.NumSheets; index++)
            {
                sheetNames.Add(this._Workbook.getSheetName(index));
            }
            return sheetNames;
        }

        /// <summary>
        /// 
        /// </summary>
        internal WorkBook Workbook
        {
            get { return this._Workbook; }
        }

        /// <summary>
        /// Try to add a worksheet with the given name.
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public bool AddSheet(string sheetName)
        {
            int lastPos = this._Workbook.NumSheets;
            this._Workbook.insertSheets(lastPos, 1);
            this._Workbook.setSheetName(lastPos, sheetName);
            return true;
        }

        /// <summary>
        /// removes a worksheet which matches with the given sheet name.
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public bool RemoveSheet(string sheetName)
        {
            int index = this._Workbook.findSheetByName(sheetName);
            if (index != -1)
            {
                this._Workbook.deleteSheets(index, 1);
            }
            return true;
        }

        /// <summary>
        /// Saves the calling WorkBook into stream. 
        /// </summary>
        /// <param name="format">the workbook format to save.</param>
        /// <returns>return a stream wich contains the calling workbook.</returns>
        public Stream Save(ExcelFormat format)
        {
            MemoryStream stream = new MemoryStream();
            switch (format)
            {
                case ExcelFormat.Xls:
                    {
                        this._Workbook.write(stream);
                        break;
                    }
                case ExcelFormat.Xlsx:
                    {
                        this._Workbook.writeXLSX(stream);
                        break;
                    }
                default:
                    {
                        throw new NotImplementedException("Workbook format not implemented.");
                    }
            }
            return stream;
        }

        /// <summary>
        /// Saves the calling WorkBook into CSV stream.
        /// </summary>
        /// <param name="sheetname">the sheet name to choose for saving into output stream.</param>
        /// <param name="separator">the separator to use into output.</param>
        /// <param name="unicode">indicates that all chars will be save with unicode encoding.</param>
        /// <param name="keepCellsFormat">indicates that values will be saved with the current value formats.</param>
        /// <returns>return a stream wich contains the calling workbook.</returns>
        public Stream SaveIntoCSV(string sheetname, char separator, bool unicode, bool keepCellsFormat)
        {
            if (string.IsNullOrEmpty(sheetname))
                throw new ArgumentException("the worksheet with the given sheetname cannot be null or empty.");

            Stream stream = null;
            try
            {
                this._Workbook.Lock();
                stream = TransformIntoCSV(this._Workbook.findSheetByName(sheetname), separator, unicode, keepCellsFormat);
            }
            finally
            {
                this._Workbook.unLock();
            }
            return stream;
        }

        /// <summary>
        /// Saves all sheets from the calling WorkBook into CSV streams.
        /// </summary>
        /// <param name="separator">the separator to use into output.</param>
        /// <param name="unicode">indicates that all chars will be save with unicode encoding.</param>
        /// <param name="keepCellsFormat">indicates that values will be saved with the current value formats.</param>
        /// <returns>return a disctionary which contains streams associated with own sheetnames.</returns>
        public IDictionary<string, Stream> SaveIntoCSV(char separator, bool unicode, bool keepCellsFormat)
        {
            Dictionary<string, Stream> output = new Dictionary<string, Stream>();
            int sheetCounter = this._Workbook.NumSheets;

            Stream stream = null;
            string sheetname = null;

            try
            {
                this._Workbook.Lock();
                for (int ind = 0; ind < sheetCounter; ind++)
                {
                    sheetname = this._Workbook.getSheetName(ind);
                    stream = TransformIntoCSV(ind, separator, unicode, keepCellsFormat);

                    output.Add(sheetname, stream);
                }
            }
            finally
            {
                this._Workbook.unLock();
            }
            return output;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheetIndex"></param>
        /// <param name="separator"></param>
        /// <param name="unicode"></param>
        /// <param name="keepCellsFormat"></param>
        /// <returns></returns>
        private Stream TransformIntoCSV(int sheetIndex, char separator, bool unicode, bool keepCellsFormat)
        {
            MemoryStream stream = null;
            if (sheetIndex != -1)
            {
                this._Workbook.Sheet = sheetIndex;
                stream = new MemoryStream();
                this._Workbook.writeCSV(stream, separator, unicode, !keepCellsFormat);
            }
            else
            {
                throw new ArgumentException("The worksheet with the given name doesn't exists.");
            }
            return stream;
        }
    }
}

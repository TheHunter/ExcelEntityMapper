using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ExcelEntityMapper;
using ExcelEntityMapper.Impl.BIFF;
using ExcelEntityMapper.Impl.Xml;
using NUnit.Framework;
using ClosedXML.Excel;
using TestExcelEntityMapper.Domain;

namespace ExcelEntityMapperTest
{
    /// <summary>
    /// 
    /// </summary>
    public class SheetXmlFilteredTest
        : SheetFilteredTest
    {
        private byte[] resourceXml;
        private byte[] emptyResourceXml;

        /// <summary>
        /// 
        /// </summary>
        public override void OnStartUp()
        {
            base.OnStartUp();
            this.resourceXml = Properties.Resources.input_XLSX_;
            this.emptyResourceXml = Properties.Resources.empty_XLSX;
        }

        [Test]
        [Category("ReaderXLSX")]
        [Description("A test which demostrate how can be read a xlsx file.")]
        public void ReadObjectsTest1()
        {
            XLSheetFilteredMapper<Person> sheet = new XLSheetFilteredMapper<Person>(1, true, this.PropertyMappersPerson);
            sheet.SheetName = "Persons";
            sheet.BeforeMapping = n => n.OwnCar = new Car();

            IXLWorkBook workbook = new XLWorkBook(new MemoryStream(this.resourceXml));

            sheet.InjectWorkBook(workbook);

            Dictionary<int, Person> buffer = new Dictionary<int, Person>();
            sheet.ReadObjects(buffer);

            Assert.IsTrue(buffer.Any());
        }

        [Test]
        [Category("ReaderXLSX")]
        [Description("Reading a empty workbook without header.")]
        public void ReadObjectsTest2()
        {
            IXLSheetFiltered<Person> sheet = new XLSheetFilteredMapper<Person>(1, false, this.PropertyMappersPerson);
            sheet.SheetName = "Persons2";

            IXLWorkBook workbook = new XLWorkBook(this.emptyResourceXml);

            sheet.InjectWorkBook(workbook);

            Dictionary<int, Person> buffer = new Dictionary<int, Person>();
            var count = sheet.ReadObjects(buffer); 

            Assert.IsTrue(!buffer.Any() && count == 0);
        }

        [Test]
        [Category("ReaderXLSX")]
        [Description("Reading a empty workbook with header.")]
        public void ReadObjectsTest3()
        {
            IXLSheetFiltered<Person> sheet = new XLSheetFilteredMapper<Person>(1, true, this.PropertyMappersPerson);
            sheet.SheetName = "Persons";

            IXLWorkBook workbook = new XLWorkBook(this.emptyResourceXml);
            workbook.AddSheet(sheet.SheetName);

            sheet.InjectWorkBook(workbook);

            Dictionary<int, Person> buffer = new Dictionary<int, Person>();
            var count = sheet.ReadObjects(buffer);      // nullreference exception..

            Assert.IsTrue(!buffer.Any() && count == 0);
        }

        [Test]
        [Category("WriterXLSX")]
        [Description("A test which demostrate how can be saved a new workbook with header of properties mapped.")]
        public void WriteObjectsTest1()
        {
            IXLSheetFiltered<Person> sheet = new XLSheetFilteredMapper<Person>(1, true, this.PropertyMappersPerson);
            sheet.SheetName = "Persons";

            XLWorkBook workbook = new XLWorkBook();
            workbook.AddSheet("Persons");

            sheet.InjectWorkBook(workbook);

            var count = sheet.WriteObjects(GetDefaultPersons());
            Assert.IsTrue(count > 0);

            WriteFileFromStream(Path.Combine(this.OutputPath, "test_output_header.xlsx"), workbook.Save());
        }

        [Test]
        [Category("WriterXLSX")]
        [Description("A test which demostrate how can be saved a new workbook without header of properties mapped.")]
        public void WriteObjectsTest2()
        {
            IXLSheetFiltered<Person> sheet = new XLSheetFilteredMapper<Person>(1, false, this.PropertyMappersPerson);
            sheet.SheetName = "Persons";

            XLWorkBook workbook = new XLWorkBook();
            workbook.AddSheet("Persons");

            sheet.InjectWorkBook(workbook);

            var count = sheet.WriteObjects(GetDefaultPersons());
            Assert.IsTrue(count > 0);

            WriteFileFromStream(Path.Combine(this.OutputPath, "test_output_noHeader.xlsx"), workbook.Save());
        }
    }
}

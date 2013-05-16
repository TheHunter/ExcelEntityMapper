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
            IXLSheetFiltered<Person> sheet = new XLSheetFilteredMapper<Person>(1, this.PropertyMappersPerson);
            sheet.SheetName = "Persons";
            sheet.BeforeReading = n => n.OwnCar = new Car();

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
            IXLSheetFiltered<Person> sheet = new XLSheetFilteredMapper<Person>(0, this.PropertyMappersPerson);
            sheet.SheetName = "Persons3";
            sheet.BeforeReading = n => n.OwnCar = new Car();

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
            IXLSheetFiltered<Person> sheet = new XLSheetFilteredMapper<Person>(1, this.PropertyMappersPerson);
            sheet.SheetName = "Persons";
            sheet.BeforeReading = n => n.OwnCar = new Car();

            IXLWorkBook workbook = new XLWorkBook(this.emptyResourceXml);
            workbook.AddSheet(sheet.SheetName);

            sheet.InjectWorkBook(workbook);

            Dictionary<int, Person> buffer = new Dictionary<int, Person>();
            var count = sheet.ReadObjects(buffer);
        }

        [Test]
        [Category("WriterXLSX")]
        [Description("A test which demostrate how can be saved a new workbook with header of properties mapped.")]
        public void WriteObjectsTest1()
        {
            IXLSheetFiltered<Person> sheet = new XLSheetFilteredMapper<Person>(1, this.PropertyMappersPerson);
            sheet.SheetName = "Persons";
            sheet.BeforeReading = n => n.OwnCar = new Car();

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
            IXLSheetFiltered<Person> sheet = new XLSheetFilteredMapper<Person>(0, this.PropertyMappersPerson);
            sheet.SheetName = "Persons";
            sheet.BeforeReading = n => n.OwnCar = new Car();

            XLWorkBook workbook = new XLWorkBook();
            workbook.AddSheet("Persons");

            sheet.InjectWorkBook(workbook);

            var count = sheet.WriteObjects(GetDefaultPersons());
            Assert.IsTrue(count > 0);

            WriteFileFromStream(Path.Combine(this.OutputPath, "test_output_noHeader.xlsx"), workbook.Save());
        }

        [Test]
        [Category("WriterXLSX")]
        [Description("A test which demostrate how can be saved a new workbook with header of properties mapped, using a sheet with header.")]
        public void WriteObjectsTest3()
        {
            IXLSheetFiltered<Person> sheet = new XLSheetFilteredMapper<Person>(1, this.PropertyMappersPerson);
            sheet.SheetName = "Persons";
            sheet.BeforeReading = n => n.OwnCar = new Car();

            IXLWorkBook workbook = new XLWorkBook(this.emptyResourceXml);
            workbook.AddSheet(sheet.SheetName);

            sheet.InjectWorkBook(workbook);

            Dictionary<int, Person> buffer = new Dictionary<int, Person>();
            var count = sheet.WriteObjects(GetDefaultPersons());

            Assert.IsTrue(count > 0);
            WriteFileFromStream(Path.Combine(this.OutputPath, "test_output_Header2.xlsx"), workbook.Save());
        }

        [Test]
        [Category("WriterXLSX")]
        [Description("A test which demostrate how can be written two typed objects into the same sheet.")]
        public void WriteObjectsTest4()
        {
            IXLWorkBook workbook = new XLWorkBook(this.emptyResourceXml);

            // a sheet mapper for saving objects..
            IXLSheetFiltered<Person> sheet = new XLSheetFilteredMapper<Person>(1, this.PropertyMappersPerson);
            sheet.SheetName = "Persons";
            sheet.BeforeReading = n => n.OwnCar = new Car();
            sheet.InjectWorkBook(workbook);

            // objects will be written into "Persons" sheet
            var personsWritten = sheet.WriteObjects(GetDefaultPersons());
            Assert.IsTrue(personsWritten > 0);

            // then, the same objects will be written into another sheet ("Persons2")
            sheet.SheetName = "Persons2";
            var personsWritten2 = sheet.WriteObjects(GetDefaultPersons());
            Assert.IsTrue(personsWritten2 > 0);


            // a sheet mapper for saving objects..
            IXLSheetFiltered<Car> sheetCar = new XLSheetFilteredMapper<Car>(1, GetCarMapper1());
            sheetCar.SheetName = "Persons";
            sheetCar.InjectWorkBook(workbook);

            var carsWritten = sheetCar.WriteObjects(GetDefaultCars());
            Assert.IsTrue(carsWritten > 0);

            sheetCar.SheetName = "Persons2";
            var carsWritten2 = sheetCar.WriteObjects(GetDefaultCars());
            Assert.IsTrue(carsWritten2 > 0);

            WriteFileFromStream(Path.Combine(this.OutputPath, "test_output_Header3.xlsx"), workbook.Save());

        }
    }
}

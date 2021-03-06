﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ExcelEntityMapper;
using ExcelEntityMapper.Exceptions;
using ExcelEntityMapper.Impl;
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
    public class SheetXmlWorkerTest
        : SheetWorkerTest
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
            IXLSheetFiltered<Person> sheet = new XLSheetFilteredMapper<Person>(1, this.PropertyMappersPerson)
                {
                    SheetName = "Persons",
                    BeforeReading = person => person.OwnCar = new Car()
                };

            IXLWorkBook workbook = new XLWorkBook(new MemoryStream(this.resourceXml));

            sheet.InjectWorkBook(workbook);

            Dictionary<int, Person> buffer = new Dictionary<int, Person>();
            var counter = sheet.ReadObjects(buffer);

            Assert.IsTrue(counter == 4);
        }

        [Test]
        [Category("ReaderXLSX")]
        [Description("Reading a empty worksheet without header.")]
        public void ReadObjectsTest2()
        {
            IXLSheetFiltered<Person> sheet = new XLSheetFilteredMapper<Person>(0, this.PropertyMappersPerson);
            sheet.SheetName = "Persons3";
            sheet.BeforeReading = n => n.OwnCar = new Car();

            IXLWorkBook workbook = new XLWorkBook(this.emptyResourceXml);

            sheet.InjectWorkBook(workbook);

            Dictionary<int, Person> buffer = new Dictionary<int, Person>();
            var counter = sheet.ReadObjects(buffer);

            Assert.IsTrue(!buffer.Any() && counter == 0);
        }

        [Test]
        [Category("ReaderXLSX")]
        [Description("Reading a empty worksheet with header.")]
        public void ReadObjectsTest3()
        {
            IXLSheetFiltered<Person> sheet = new XLSheetFilteredMapper<Person>(1, this.PropertyMappersPerson);
            sheet.SheetName = "Persons4";
            sheet.BeforeReading = n => n.OwnCar = new Car();

            IXLWorkBook workbook = new XLWorkBook(this.emptyResourceXml);
            workbook.AddSheet(sheet.SheetName);

            sheet.InjectWorkBook(workbook);

            Dictionary<int, Person> buffer = new Dictionary<int, Person>();
            var counter = sheet.ReadObjects(buffer);

            Assert.IsTrue(!buffer.Any() && counter == 0);
        }

        [Test]
        [Category("ReaderXLSX")]
        [Description("Finding the row index of the first readable row.")]
        public void FindFirstIndexRow()
        {
            IXLSheetFiltered<Person> sheet = new XLSheetFilteredMapper<Person>(1, this.PropertyMappersPerson);
            IXLWorkBook workbook = new XLWorkBook(this.emptyResourceXml);
            sheet.InjectWorkBook(workbook);

            IXLSheetFiltered<Person> sheetNoHeader = new XLSheetFilteredMapper<Person>(this.PropertyMappersPerson);
            IXLWorkBook wb = new XLWorkBook(this.emptyResourceXml);
            sheetNoHeader.InjectWorkBook(wb);

            Assert.AreEqual(sheet.GetIndexFirstRow("Persons"), 5);
            Assert.AreEqual(sheet.GetIndexFirstRow("Persons4"), -1);

            Assert.AreEqual(sheetNoHeader.GetIndexFirstRow("Persons2"), 8);
            Assert.AreEqual(sheetNoHeader.GetIndexFirstRow("Persons3"), -1);
        }

        [Test]
        [Category("ReaderXLSX")]
        [Description("Finding the row index of the first readable row.")]
        public void ReadRowOnIndex()
        {
            IXLSheetFiltered<Person> sheet = new XLSheetFilteredMapper<Person>(1, this.PropertyMappersPerson);

            IXLWorkBook workbook = new XLWorkBook(this.emptyResourceXml);
            sheet.InjectWorkBook(workbook);

            //the return object must be validated because It could be initialized with wrongs values from worksheet source.
            Assert.IsNull(sheet.ReadObject("Persons4", 1));
            Assert.IsNotNull(sheet.ReadObject("Persons", 5));
        }

        [Test]
        [Category("ReaderXLSX")]
        [Description("Reading a worksheet with filter.")]
        public void ReadFilteredObjects()
        {
            IXLSheetFiltered<Person> sheet = new XLSheetFilteredMapper<Person>(1, this.PropertyMappersPerson);
            sheet.SheetName = "Persons";
            sheet.BeforeReading = n => n.OwnCar = new Car();

            IXLWorkBook workbook = new XLWorkBook(this.emptyResourceXml);
            workbook.AddSheet(sheet.SheetName);

            sheet.InjectWorkBook(workbook);

            Dictionary<int, Person> buffer = new Dictionary<int, Person>();
            sheet.ReadFilteredObjects(buffer, person => person.MarriedYear > 1999);
            Assert.IsTrue(buffer.Count > 0);
            buffer.Clear();
            
            sheet.ReadFilteredObjects(buffer, n => n.MarriedYear > 2000);
            Assert.IsTrue(buffer.Count == 0);
        }

        [Test]
        [Category("WrongReadOperation")]
        [Description("It throws an exception because SheetWorker wasn't injected the workbook to read.")]
        [ExpectedException(typeof(UnReadableSheetException))]
        public void WrongReadObjectsTest1()
        {
            IXLSheetFiltered<Person> sheet = new XLSheetFilteredMapper<Person>(1, this.PropertyMappersPerson);
            sheet.SheetName = "Persons";
            sheet.BeforeReading = n => n.OwnCar = new Car();

            sheet.ReadObjects(new Dictionary<int, Person>());
        }

        [Test]
        [Category("FinderIndex")]
        [Description("It tries to find the last used index row.")]
        public void FindLastIndexRowTest1()
        {
            IXLSheetFiltered<Person> sheetHeader = new XLSheetFilteredMapper<Person>(1, this.PropertyMappersPerson);
            IXLWorkBook workbook = new XLWorkBook(this.emptyResourceXml);
            sheetHeader.InjectWorkBook(workbook);

            IXLSheetFiltered<Person> sheetNoHeader = new XLSheetFilteredMapper<Person>(this.PropertyMappersPerson);
            IXLWorkBook wb = new XLWorkBook(this.emptyResourceXml);
            sheetNoHeader.InjectWorkBook(wb);

            Assert.AreEqual(sheetHeader.GetIndexLastRow("Persons"), 6);
            Assert.AreEqual(sheetHeader.GetIndexLastRow("Persons4"), 6);

            Assert.AreEqual(sheetNoHeader.GetIndexLastRow("Persons2"), 9);
            Assert.AreEqual(sheetNoHeader.GetIndexLastRow("Persons3"), 0);
        }

        [Test]
        [Category("WriterXLSX")]
        [Description("A test which demostrate how can be saved a new worksheet with header of properties mapped.")]
        public void WriteObjectsTest1()
        {
            IXLSheetFiltered<Person> sheet = new XLSheetFilteredMapper<Person>(1, this.PropertyMappersPerson);
            sheet.SheetName = "Persons";

            XLWorkBook workbook = new XLWorkBook();
            workbook.AddSheet(sheet.SheetName);

            sheet.InjectWorkBook(workbook);

            var count = sheet.WriteObjects(GetDefaultPersons());
            Assert.IsTrue(count > 0);

            WriteFileFromStream(Path.Combine(this.OutputPath, "test_output_header.xlsx"), workbook.Save());
        }

        [Test]
        [Category("WriterXLSX")]
        [Description("A test which demostrate how can be saved a new worksheet without header of properties mapped.")]
        public void WriteObjectsTest2()
        {
            IXLSheetFiltered<Person> sheet = new XLSheetFilteredMapper<Person>(0, this.PropertyMappersPerson);
            sheet.SheetName = "Persons";

            XLWorkBook workbook = new XLWorkBook();
            workbook.AddSheet("Persons");

            sheet.InjectWorkBook(workbook);

            var count = sheet.WriteObjects(GetDefaultPersons());
            Assert.IsTrue(count > 0);

            WriteFileFromStream(Path.Combine(this.OutputPath, "test_output_noHeader.xlsx"), workbook.Save());
        }

        [Test]
        [Category("WriterXLSX")]
        [Description("A test which demostrate how can be saved a new worksheet with header of properties mapped.")]
        public void WriteObjectsTest3()
        {
            IXLSheetFiltered<Person> sheet = new XLSheetFilteredMapper<Person>(1, this.PropertyMappersPerson);
            sheet.SheetName = "Persons";

            IXLWorkBook workbook = new XLWorkBook(this.emptyResourceXml);
            workbook.AddSheet(sheet.SheetName);

            sheet.InjectWorkBook(workbook);

            var lista = new List<Person>(GetDefaultPersons());
            //a null object which will not be saved into worksheet.
            lista.Insert(1, null);

            var count = sheet.WriteObjects(lista);

            Assert.AreEqual(count, lista.Count - 1);
            WriteFileFromStream(Path.Combine(this.OutputPath, "test_output_Header2.xlsx"), workbook.Save());
        }

        [Test]
        [Category("WriterXLS")]
        [Description("A test which demostrate how can be appended instances on worksheet without header of properties mapped.")]
        public void WriteObjectsTest4()
        {
            IXLSheetFiltered<Person> sheet = new XLSheetFilteredMapper<Person>(this.PropertyMappersPerson);
            IXLWorkBook workbook = new XLWorkBook(this.emptyResourceXml);

            sheet.InjectWorkBook(workbook);

            var count1 = sheet.WriteObjects("Persons2", GetDefaultPersons());
            var count2 = sheet.WriteObjects("Persons3", GetDefaultPersons());

            Assert.IsTrue(count1 > 0 && count2 > 0);
            WriteFileFromStream(Path.Combine(this.OutputPath, "test_output_NoHeader2.xlsx"), workbook.Save());
        }

        [Test]
        [Category("WriterXLSX")]
        [Description("")]
        public void WriteObjectsByIndexTest()
        {
            IXLSheetFiltered<Person> sheetNoHeader = new XLSheetFilteredMapper<Person>(this.PropertyMappersPerson);
            IXLSheetFiltered<Person> sheetHeader = new XLSheetFilteredMapper<Person>(this.PropertyMappersPerson);

            IXLWorkBook workbook = new XLWorkBook(this.emptyResourceXml);

            sheetNoHeader.InjectWorkBook(workbook);
            sheetHeader.InjectWorkBook(workbook);

            var p1 = sheetNoHeader.GetIndexLastRow("Persons");
            var p2 = sheetNoHeader.GetIndexLastRow("Persons2");
            var p3 = sheetNoHeader.GetIndexLastRow("Persons3");
            var p4 = sheetNoHeader.GetIndexLastRow("Persons4");

            var persons = GetDefaultPersons();
            foreach (var person in persons)
            {
                sheetHeader.WriteObject("Persons", ++p1, person);
                sheetNoHeader.WriteObject("Persons2", ++p2, person);
                sheetNoHeader.WriteObject("Persons3", ++p3, person);
                sheetHeader.WriteObject("Persons4", ++p4, person);
            }
            WriteFileFromStream(Path.Combine(this.OutputPath, "test_output_byIndex.xlsx"), workbook.Save());
        }

        [Test]
        [Category("WriterXLSX")]
        [Description("")]
        public void AppendObjectsTest()
        {
            IXLSheetFiltered<Person> sheetNoHeader = new XLSheetFilteredMapper<Person>(this.PropertyMappersPerson);
            IXLSheetFiltered<Person> sheetHeader = new XLSheetFilteredMapper<Person>(this.PropertyMappersPerson);

            IXLWorkBook workbook = new XLWorkBook(this.emptyResourceXml);

            sheetNoHeader.InjectWorkBook(workbook);
            sheetHeader.InjectWorkBook(workbook);

            var persons = GetDefaultPersons();
            foreach (var person in persons)
            {
                sheetHeader.WriteObject("Persons", person);
                sheetNoHeader.WriteObject("Persons2", person);
                sheetNoHeader.WriteObject("Persons3", person);
                sheetHeader.WriteObject("Persons4", person);
            }
            WriteFileFromStream(Path.Combine(this.OutputPath, "test_output_AppendObjects.xlsx"), workbook.Save());
        }

        [Test]
        [Category("WriterXLSX")]
        [Description("A test which demostrate how can be written two typed objects into the same sheet.")]
        public void WriteObjectsTest5()
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

        [Test]
        [Category("WrongWorkBook")]
        [Description("Injecting a wrong WorkBook")]
        [ExpectedException(typeof(WorkBookException))]
        public void WrongWorkBookTest1()
        {
            IXLSheetFiltered<Person> sheet = new XLSheetFilteredMapper<Person>(1, this.PropertyMappersPerson);
            // null workbook is wrong argument
            sheet.InjectWorkBook(null);
        }

        [Test]
        [Category("WrongWorkBook")]
        [Description("Injecting a wrong WorkBook")]
        [ExpectedException(typeof(SheetParameterException))]
        public void WrongWorkBookTest2()
        {
            IXLSheetFiltered<Person> sheet = new XLSheetFilteredMapper<Person>(1, this.PropertyMappersPerson);
            // Workbook instance must be compatible with the mapper to associate.
            sheet.InjectWorkBook(new XWorkBook());
        }

        [Test]
        [Category("SheetConverter")]
        [Description("A SheetMapper which makes a new IXLSheetWorker into XML SheetMapper and Biff SheetMapper.")]
        public void SheetMapperConverterTest()
        {
            IXLSheetFiltered<Person> sheet = new XLSheetFilteredMapper<Person>(0, this.PropertyMappersPerson);
            var sheetWorker1 = sheet.AsXmlSheetWorker();
            Assert.IsNotNull(sheetWorker1);

            var sheetWorker2 = sheet.AsBiffSheetWorker();
            Assert.IsNotNull(sheetWorker2);
        }

        [Test]
        [Category("SheetConverter")]
        [Description("A SheetWriter which makes a new IXLSheetWorker into XML SheetReader and Biff SheetReader.")]
        public void SheetWriterConverterTest()
        {
            IXLSheetFiltered<Person> sheet = new XLSheetFilteredMapper<Person>(0, this.PropertyMappersPerson);
            var sheetWorker1 = sheet.AsXmlReader();
            Assert.IsNotNull(sheetWorker1);

            var sheetWorker2 = sheet.AsBiffReader();
            Assert.IsNotNull(sheetWorker2);
        }

        [Test]
        [Category("SheetConverter")]
        [Description("A SheetWriter which makes a new IXLSheetWorker into XML SheetReader and Biff SheetWriter.")]
        public void SheetReaderConverterTest()
        {
            IXLSheetFiltered<Person> sheet = new XLSheetFilteredMapper<Person>(0, this.PropertyMappersPerson);
            var sheetWorker1 = sheet.AsXmlWriter();
            Assert.IsNotNull(sheetWorker1);

            var sheetWorker2 = sheet.AsBiffWriter();
            Assert.IsNotNull(sheetWorker2);
        }
    }
}

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
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NUnit.Framework;
using System.Data;
using TestExcelEntityMapper.Domain;

namespace ExcelEntityMapperTest
{
    /// <summary>
    /// 
    /// </summary>
    public class SheetBiffWorkerTest
        : SheetWorkerTest
    {
        private byte[] resource;
        private byte[] emptyResource;

        /// <summary>
        /// 
        /// </summary>
        public override void OnStartUp()
        {
            base.OnStartUp();
            this.resource = Properties.Resources.input_XLS_;
            this.emptyResource = Properties.Resources.empty_XLS;
        }

        [Test]
        [Category("ReaderXLS")]
        [Description("A test which demostrate how can be read a xls file.")]
        public void ReadObjectsTest1()
        {
            IXLSheetFiltered<Person> sheet = new XSheetFilteredMapper<Person>(1, this.PropertyMappersPerson)
                {
                    SheetName = "Persons",
                    BeforeReading = person => person.OwnCar = new Car()
                };

            IXLWorkBook workbook = new XWorkBook(new MemoryStream(this.resource));

            sheet.InjectWorkBook(workbook);

            Dictionary<int, Person> buffer = new Dictionary<int, Person>();
            var counter = sheet.ReadObjects(buffer);

            Assert.IsTrue(counter == 4);
        }

        [Test]
        [Category("ReaderXLS")]
        [Description("Reading a empty worksheet without header.")]
        public void ReadObjectsTest2()
        {
            IXLSheetFiltered<Person> sheet = new XSheetFilteredMapper<Person>(0, this.PropertyMappersPerson);
            sheet.SheetName = "Persons3";
            sheet.BeforeReading = n => n.OwnCar = new Car();

            IXLWorkBook workbook = new XWorkBook(this.emptyResource);

            sheet.InjectWorkBook(workbook);

            Dictionary<int, Person> buffer = new Dictionary<int, Person>();
            var counter = sheet.ReadObjects(buffer);

            Assert.IsTrue(!buffer.Any() && counter == 0); 
        }

        [Test]
        [Category("ReaderXLS")]
        [Description("Reading a empty worksheet with header.")]
        public void ReadObjectsTest3()
        {
            IXLSheetFiltered<Person> sheet = new XSheetFilteredMapper<Person>(1, this.PropertyMappersPerson);
            sheet.SheetName = "Persons4";
            sheet.BeforeReading = n => n.OwnCar = new Car();

            IXLWorkBook workbook = new XWorkBook(this.emptyResource);
            workbook.AddSheet(sheet.SheetName);

            sheet.InjectWorkBook(workbook);

            Dictionary<int, Person> buffer = new Dictionary<int, Person>();
            var counter = sheet.ReadObjects(buffer);

            Assert.IsTrue(!buffer.Any() && counter == 0);
        }

        [Test]
        [Category("ReaderXLS")]
        [Description("Finding the row index of the first readable row.")]
        public void FindFirstIndexRow()
        {
            IXLSheetFiltered<Person> sheet = new XSheetFilteredMapper<Person>(1, this.PropertyMappersPerson);
            IXLWorkBook workbook = new XWorkBook(this.emptyResource);
            sheet.InjectWorkBook(workbook);

            IXLSheetFiltered<Person> sheetNoHeader = new XSheetFilteredMapper<Person>(this.PropertyMappersPerson);
            IXLWorkBook wb = new XWorkBook(this.emptyResource);
            sheetNoHeader.InjectWorkBook(wb);

            Assert.AreEqual(sheet.GetIndexFirstRow("Persons"), 5);
            Assert.AreEqual(sheet.GetIndexFirstRow("Persons4"), -1);

            Assert.AreEqual(sheetNoHeader.GetIndexFirstRow("Persons2"), 8);
            Assert.AreEqual(sheetNoHeader.GetIndexFirstRow("Persons3"), -1);
        }

        [Test]
        [Category("ReaderXLS")]
        [Description("Finding the row index of the first readable row.")]
        public void ReadRowOnIndex()
        {
            IXLSheetFiltered<Person> sheet = new XSheetFilteredMapper<Person>(1, this.PropertyMappersPerson);

            IXLWorkBook workbook = new XWorkBook(this.emptyResource);
            sheet.InjectWorkBook(workbook);

            //the return object must be validated because It could be initialized with wrongs values from worksheet source.
            Assert.IsNull(sheet.ReadObject("Persons4", 1));
            Assert.IsNotNull(sheet.ReadObject("Persons", 5));
        }

        [Test]
        [Category("ReaderXLS")]
        [Description("Reading a worksheet with filter.")]
        public void ReadFilteredObjects()
        {
            IXLSheetFiltered<Person> sheet = new XSheetFilteredMapper<Person>(1, this.PropertyMappersPerson);
            sheet.SheetName = "Persons";
            sheet.BeforeReading = n => n.OwnCar = new Car();

            IXLWorkBook workbook = new XWorkBook(this.emptyResource);
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
            IXLSheetFiltered<Person> sheet = new XSheetFilteredMapper<Person>(1, this.PropertyMappersPerson);
            sheet.SheetName = "Persons";
            sheet.BeforeReading = n => n.OwnCar = new Car();

            sheet.ReadObjects(new Dictionary<int, Person>());
        }

        [Test]
        [Category("FinderIndex")]
        [Description("It tries to find the last used index row.")]
        public void FindLastIndexRowTest1()
        {
            IXLSheetFiltered<Person> sheetHeader = new XSheetFilteredMapper<Person>(1, this.PropertyMappersPerson);
            IXLWorkBook workbook = new XWorkBook(this.emptyResource);
            sheetHeader.InjectWorkBook(workbook);

            IXLSheetFiltered<Person> sheetNoHeader = new XSheetFilteredMapper<Person>(this.PropertyMappersPerson);
            IXLWorkBook wb = new XWorkBook(this.emptyResource);
            sheetNoHeader.InjectWorkBook(wb);

            Assert.AreEqual(sheetHeader.GetIndexLastRow("Persons"), 6);
            Assert.AreEqual(sheetHeader.GetIndexLastRow("Persons4"), 6);

            Assert.AreEqual(sheetNoHeader.GetIndexLastRow("Persons2"), 9);
            Assert.AreEqual(sheetNoHeader.GetIndexLastRow("Persons3"), 0);
        }

        [Test]
        [Category("WriterXLS")]
        [Description("A test which demostrate how can be saved a new workbook with header of properties mapped.")]
        public void WriteObjectsTest1()
        {
            IXLSheetFiltered<Person> sheet = new XSheetFilteredMapper<Person>(1, this.PropertyMappersPerson);
            sheet.SheetName = "Persons";

            XWorkBook workbook = new XWorkBook();
            workbook.AddSheet(sheet.SheetName);

            sheet.InjectWorkBook(workbook);

            var count = sheet.WriteObjects(GetDefaultPersons());
            Assert.IsTrue(count > 0);

            WriteFileFromStream(Path.Combine(this.OutputPath, "test_output_header.xls"), workbook.Save());
        }

        [Test]
        [Category("WriterXLS")]
        [Description("A test which demostrate how can be saved a new worksheet without header of properties mapped.")]
        public void WriteObjectsTest2()
        {
            IXLSheetFiltered<Person> sheet = new XSheetFilteredMapper<Person>(0, this.PropertyMappersPerson);
            sheet.SheetName = "Persons";

            XWorkBook workbook = new XWorkBook();
            workbook.AddSheet("Persons");

            sheet.InjectWorkBook(workbook);

            var count = sheet.WriteObjects(GetDefaultPersons());
            Assert.IsTrue(count > 0);

            WriteFileFromStream(Path.Combine(this.OutputPath, "test_output_noHeader.xls"), workbook.Save());
        }

        [Test]
        [Category("WriterXLS")]
        [Description("A test which demostrate how can be appended instances on worksheet with header of properties mapped.")]
        public void WriteObjectsTest3()
        {
            IXLSheetFiltered<Person> sheet = new XSheetFilteredMapper<Person>(1, this.PropertyMappersPerson);
            sheet.SheetName = "Persons";

            IXLWorkBook workbook = new XWorkBook(this.emptyResource);
            workbook.AddSheet(sheet.SheetName);

            sheet.InjectWorkBook(workbook);

            var lista = new List<Person>(GetDefaultPersons());
            //a null object which will not be saved into worksheet.
            lista.Insert(1, null);

            var count = sheet.WriteObjects(lista);

            Assert.AreEqual(count, lista.Count - 1);
            WriteFileFromStream(Path.Combine(this.OutputPath, "test_output_Header2.xls"), workbook.Save());
        }

        [Test]
        [Category("WriterXLS")]
        [Description("A test which demostrate how can be appended instances on worksheet without header of properties mapped.")]
        public void WriteObjectsTest4()
        {
            IXLSheetFiltered<Person> sheet = new XSheetFilteredMapper<Person>(this.PropertyMappersPerson);
            IXLWorkBook workbook = new XWorkBook(this.emptyResource);

            sheet.InjectWorkBook(workbook);

            var count1 = sheet.WriteObjects("Persons2", GetDefaultPersons());
            var count2 = sheet.WriteObjects("Persons3", GetDefaultPersons());

            Assert.IsTrue(count1 > 0 && count2 > 0);
            WriteFileFromStream(Path.Combine(this.OutputPath, "test_output_NoHeader2.xls"), workbook.Save());
        }

        [Test]
        [Category("WriterXLS")]
        [Description("")]
        public void WriteObjectsByIndexTest()
        {
            IXLSheetFiltered<Person> sheetNoHeader = new XSheetFilteredMapper<Person>(this.PropertyMappersPerson);
            IXLSheetFiltered<Person> sheetHeader = new XSheetFilteredMapper<Person>(this.PropertyMappersPerson);

            IXLWorkBook workbook = new XWorkBook(this.emptyResource);

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
            WriteFileFromStream(Path.Combine(this.OutputPath, "test_output_byIndex.xls"), workbook.Save());
        }

        [Test]
        [Category("WriterXLS")]
        [Description("")]
        public void AppendObjectsTest()
        {
            IXLSheetFiltered<Person> sheetNoHeader = new XSheetFilteredMapper<Person>(this.PropertyMappersPerson);
            IXLSheetFiltered<Person> sheetHeader = new XSheetFilteredMapper<Person>(this.PropertyMappersPerson);

            IXLWorkBook workbook = new XWorkBook(this.emptyResource);

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
            WriteFileFromStream(Path.Combine(this.OutputPath, "test_output_AppendObjects.xls"), workbook.Save());
        }

        [Test]
        [Category("WriterXLS")]
        [Description("A test which demostrate how can be written two typed objects into the same sheet.")]
        public void WriteObjectsTest5()
        {
            IXLWorkBook workbook = new XWorkBook(this.emptyResource);

            // a sheet mapper for saving objects..
            IXLSheetFiltered<Person> sheet = new XSheetFilteredMapper<Person>(1, this.PropertyMappersPerson);
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
            IXLSheetFiltered<Car> sheetCar = new XSheetFilteredMapper<Car>(1, GetCarMapper1());
            sheetCar.SheetName = "Persons";
            sheetCar.InjectWorkBook(workbook);

            var carsWritten = sheetCar.WriteObjects(GetDefaultCars());
            Assert.IsTrue(carsWritten > 0);

            sheetCar.SheetName = "Persons2";
            var carsWritten2 = sheetCar.WriteObjects(GetDefaultCars());
            Assert.IsTrue(carsWritten2 > 0);

            WriteFileFromStream(Path.Combine(this.OutputPath, "test_output_Header3.xls"), workbook.Save());
        }

        [Test]
        [Category("WrongWorkBook")]
        [Description("Injecting a wrong WorkBook")]
        [ExpectedException(typeof(WorkBookException))]
        public void WrongWorkBookTest1()
        {
            IXLSheetFiltered<Person> sheet = new XSheetFilteredMapper<Person>(1, this.PropertyMappersPerson);
            // null workbook is wrong argument
            sheet.InjectWorkBook(null);
        }

        [Test]
        [Category("WrongWorkBook")]
        [Description("Injecting a wrong WorkBook")]
        [ExpectedException(typeof(SheetParameterException))]
        public void WrongWorkBookTest2()
        {
            IXLSheetFiltered<Person> sheet = new XSheetFilteredMapper<Person>(1, this.PropertyMappersPerson);
            // Workbook instance must be compatible with the mapper to associate.
            sheet.InjectWorkBook(new XLWorkBook());
        }

        [Test]
        [Category("SheetConverter")]
        [Description("A SheetMapper which makes a new IXLSheetWorker into XML SheetMapper and Biff SheetMapper.")]
        public void SheetMapperConverterTest()
        {
            IXLSheetFiltered<Person> sheet = new XSheetFilteredMapper<Person>(0, this.PropertyMappersPerson);
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
            IXLSheetFiltered<Person> sheet = new XSheetFilteredMapper<Person>(0, this.PropertyMappersPerson);
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
            IXLSheetFiltered<Person> sheet = new XSheetFilteredMapper<Person>(0, this.PropertyMappersPerson);
            var sheetWorker1 = sheet.AsXmlWriter();
            Assert.IsNotNull(sheetWorker1);

            var sheetWorker2 = sheet.AsBiffWriter();
            Assert.IsNotNull(sheetWorker2);
        }
    }
}

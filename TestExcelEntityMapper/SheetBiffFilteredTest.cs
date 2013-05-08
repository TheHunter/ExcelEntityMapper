using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ExcelEntityMapper;
using ExcelEntityMapper.Impl.BIFF;
using ExcelEntityMapper.Impl.Xml;
using ExcelLibrary.SpreadSheet;
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
    public class SheetBiffFilteredTest
        : SheetFilteredTest
    {
        private byte[] resource;

        /// <summary>
        /// 
        /// </summary>
        public override void OnStartUp()
        {
            base.OnStartUp();
            this.resource = Properties.Resources.output_XLS_;
        }

        [Test]
        [Category("ReaderXLS")]
        [Description("A test which demostrate how can be read a xls file.")]
        public void ReadObjectsTest1()
        {
            IXLSheetFiltered<Person> sheet = new XSheetFilteredMapper<Person>(1, true, this.PropertyMappersPerson);
            sheet.SheetName = "Persons";

            IXLWorkBook workbook = new XWorkBook(new MemoryStream(this.resource));

            sheet.InjectWorkBook(workbook);

            Dictionary<int, Person> buffer = new Dictionary<int, Person>();
            sheet.ReadObjects(buffer);

            Assert.IsTrue(buffer.Any());
        }

        [Test]
        [Category("WriterXLS")]
        [Description("A test which demostrate how can be saved a new workbook with header of properties mapped.")]
        public void WriteObjectsTest1()
        {
            IXLSheetFiltered<Person> sheet = new XSheetFilteredMapper<Person>(1, true, this.PropertyMappersPerson);
            sheet.SheetName = "Persons";

            XWorkBook workbook = new XWorkBook();
            workbook.AddSheet("Persons");

            sheet.InjectWorkBook(workbook);

            var count = sheet.WriteObjects(GetDefaultPersons());
            Assert.IsTrue(count > 0);

            WriteFileFromStream(Path.Combine(this.OutputPath, "test_output_header.xls"), workbook.Save());
        }

        [Test]
        [Category("WriterXLS")]
        [Description("A test which demostrate how can be saved a new workbook without header of properties mapped.")]
        public void WriteObjectsTest2()
        {
            IXLSheetFiltered<Person> sheet = new XSheetFilteredMapper<Person>(1, false, this.PropertyMappersPerson);
            sheet.SheetName = "Persons";

            XWorkBook workbook = new XWorkBook();
            workbook.AddSheet("Persons");

            sheet.InjectWorkBook(workbook);

            var count = sheet.WriteObjects(GetDefaultPersons());
            Assert.IsTrue(count > 0);

            WriteFileFromStream(Path.Combine(this.OutputPath, "test_output_noHeader.xls"), workbook.Save());
        }

        [Test]
        [Category("Demo")]
        public void Test1()
        {
            XWorkBook workbook = new XWorkBook();
            workbook.AddSheet("newone");

            WriteFileFromStream(Path.Combine(this.OutputPath, "test_output_test.xls"), workbook.Save());
        }

        [Test]
        [Category("Demo")]
        public void Test2()
        {
            HSSFWorkbook a = new HSSFWorkbook();

            // imp of AddSheet()
            a.CreateSheet("Persons");
            a.CreateSheet("Agencies");
            

            // imp of ExistsWorkSheet()
            Assert.IsTrue(a.GetSheetIndex("Persons") != -1);
            Assert.IsTrue(a.GetSheetIndex("persons") != -1);
            Assert.IsTrue(a.GetSheetIndex("Persons_") == -1);

            // imp of GetSheetNames()
            List<string> names = new List<string>();
            for (int i = 0; i < a.NumberOfSheets; i++)
            {
                names.Add(a.GetSheetName(i));
            }
            Assert.IsTrue(names.Count == 2);

            // imp of RemoveSheet()
            a.RemoveSheetAt(a.GetSheetIndex("Agencies"));
            names.Remove("Agencies");
            Assert.IsTrue(a.NumberOfSheets == 1);



            MemoryStream mem = new MemoryStream();
            a.Write(mem);
            WriteFileFromStream(Path.Combine(this.OutputPath, "test_out2_.xls"), mem);
        }

        [Test]
        [Category("Demo")]
        public void Test3()
        {
            HSSFWorkbook a = new HSSFWorkbook(new MemoryStream(this.resource));
            var sheet = a.CreateSheet("test");

            var row0 = sheet.CreateRow(1);
            var header = sheet.GetRow(1);

            Assert.AreEqual(row0, header);
            Assert.AreSame(row0, header);
            Assert.IsNotNull(header);

            header.CreateCell(0).SetCellValue("header1");
            header.CreateCell(1).SetCellValue("header2");

            var row1 = sheet.CreateRow(2);
            row1.CreateCell(0).SetCellValue("r2:c1");
            row1.CreateCell(1).SetCellValue("r2:c2");

            var row2 = sheet.CreateRow(5);
            row2.CreateCell(0).SetCellValue("r5:c1");
            row2.CreateCell(1).SetCellValue("r5:c2");
            row2.CreateCell(2).SetCellValue(DateTime.Now.ToString("dd/MM/yyyy"));

            //var row3 = sheet.CreateRow(5);
            //row2.CreateCell(0).SetCellValue("r5:c1..");
            //row2.CreateCell(1).SetCellValue("r5:c2..");
            //row2.CreateCell(2).SetCellValue(DateTime.Now.ToString("dd/MM/yyyy"));

            var tst =sheet.GetRow(4);
            Assert.IsNull(tst);

            var lastRow = sheet.LastRowNum;     //da verificare il suo funzionamento...
            var firstRow = sheet.FirstRowNum;   //item questo.

            var sheetPerson = a.GetSheet("Persons");
            var lastrow = sheetPerson.LastRowNum;

            MemoryStream mem = new MemoryStream();
            a.Write(mem);
            WriteFileFromStream(Path.Combine(this.OutputPath, "test_out3_.xls"), mem);
        }

        [Test]
        [Category("Demo")]
        public void Test4()
        {
            HSSFWorkbook a = new HSSFWorkbook(File.Open(@"C:\Users\Diego\AppData\Local\Temp\test_out3__.xls", FileMode.Open, FileAccess.ReadWrite));
            var sheet = a.GetSheet("test2");

            var firstrow = sheet.FirstRowNum;
            var lastrow = sheet.LastRowNum;

            Assert.IsTrue(true);
        }
    }
}

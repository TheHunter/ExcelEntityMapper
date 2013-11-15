using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using ExcelEntityMapper;
using ExcelEntityMapper.Impl.BIFF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NUnit.Framework;

namespace ExcelEntityMapperTest
{
    //[TestFixture]
    public sealed class DemoBiffExcelFormat
    {
        private byte[] resource;
        private byte[] emptyResource;
        private string outputPath;

        [TestFixtureSetUp]
        public void OnStartUp()
        {
            this.resource = Properties.Resources.input_XLS_;
            this.emptyResource = Properties.Resources.empty_XLS;
            outputPath = Path.GetTempPath();
        }

        [TestFixtureTearDown]
        public void OnOverTest()
        {
            this.outputPath = null;
        }

        //[Test]
        [Category("Demo")]
        public void Test1()
        {
            XWorkBook workbook = new XWorkBook();
            workbook.AddSheet("newone");

            WriteFileFromStream(Path.Combine(this.outputPath, "test_output_test.xls"), workbook.Save());
        }

        //[Test]
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
            WriteFileFromStream(Path.Combine(this.outputPath, "test_out2_.xls"), mem);
        }

        //[Test]
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

            var rowN = sheet.CreateRow(20);
            Assert.IsNotNull(rowN);
            Assert.IsNotNull(sheet.GetRow(20));
            Assert.IsNull(sheet.GetRow(19));

            var cell20 = header.CreateCell(20);
            Console.WriteLine(cell20.CellType);

            header.CreateCell(0).SetCellValue("header1");
            header.CreateCell(1).SetCellValue("header2");

            var row1 = sheet.CreateRow(2);
            row1.CreateCell(0).SetCellValue("r2:c1");
            row1.CreateCell(1).SetCellValue("r2:c2");

            var row2 = sheet.CreateRow(5);
            row2.CreateCell(0).SetCellValue("r5:c1");
            row2.CreateCell(1).SetCellValue("r5:c2");
            row2.CreateCell(2).SetCellValue(DateTime.Now.ToString("dd/MM/yyyy"));

            // bool value
            row2.CreateCell(3).SetCellValue(true);

            // viene applicato la data corrente
            row2.CreateCell(4).SetCellValue(DateTime.Now);

            row2.CreateCell(5).SetCellValue(((byte)0));
            row2.CreateCell(6).SetCellValue(((short)124));
            row2.CreateCell(7).SetCellValue(((int)254));
            row2.CreateCell(8).SetCellValue(1000L);
            row2.CreateCell(9).SetCellValue(1000.454D);
            row2.CreateCell(10).SetCellValue("ciao a tutti \r\n... seconda riga\r\n... terza riga");

            ICell cella = row2.CreateCell(11);
            //object value = XLEntityHelper.NormalizeXlsCellValue(DateTime.Now);

            var tst = sheet.GetRow(4);
            Assert.IsNull(tst);

            var lastRow = sheet.LastRowNum;     //da verificare il suo funzionamento...
            var firstRow = sheet.FirstRowNum;   //item questo.

            var sheetPerson = a.GetSheet("Persons");
            var lastrow = sheetPerson.LastRowNum;

            MemoryStream mem = new MemoryStream();
            a.Write(mem);
            WriteFileFromStream(Path.Combine(this.outputPath, "test_out3_.xls"), mem);
        }

        //[Test]
        [Category("Demo")]
        public void Test4()
        {
            HSSFWorkbook a = new HSSFWorkbook(new MemoryStream(this.resource));
            var sheet = a.CreateSheet("test");

            var row0 = sheet.CreateRow(1);
            ICell cell1 = row0.CreateCell(0);
            ICell cell2 = row0.CreateCell(1);

            ParameterExpression instanceInvoker = Expression.Parameter(typeof(ICell), "cell");
            ParameterExpression parameter1 = Expression.Parameter(typeof(double), "value");
            MethodInfo method = typeof(ICell).GetMethod("SetCellValue", new[] { typeof(double) });

            Delegate act = Expression.Lambda<Action<ICell, double>>
                                            (Expression.Call(instanceInvoker, method, parameter1), instanceInvoker, parameter1).Compile();
            //act.Invoke(cell1, 5);
            act.DynamicInvoke(cell1, 5);
            
            Assert.AreEqual(cell1.NumericCellValue, 5);
        }

        //[Test]
        public void TestXlSTypes()
        {
            Assert.AreEqual(XLEntityHelper.NormalizeXlsCellValue(DateTime.Now).GetType(), typeof(string));
            Assert.AreEqual(XLEntityHelper.NormalizeXlsCellValue((DateTime?)DateTime.Now).GetType(), typeof(string));

            Assert.AreEqual(XLEntityHelper.NormalizeXlsCellValue(true).GetType(), typeof(string));
            Assert.AreEqual(XLEntityHelper.NormalizeXlsCellValue((bool?)true).GetType(), typeof(string));

            Assert.AreEqual(XLEntityHelper.NormalizeXlsCellValue(new StringBuilder()).GetType(), typeof(string));

            Assert.AreEqual(XLEntityHelper.NormalizeXlsCellValue((short)1).GetType(), typeof(short));
            Assert.AreEqual(XLEntityHelper.NormalizeXlsCellValue((short?)1).GetType(), typeof(short));

            Assert.AreEqual(XLEntityHelper.NormalizeXlsCellValue((byte)5).GetType(), typeof(byte));
            Assert.AreEqual(XLEntityHelper.NormalizeXlsCellValue((byte?)5).GetType(), typeof(byte));
            
            Assert.AreEqual(XLEntityHelper.NormalizeXlsCellValue(5).GetType(), typeof(int));
            Assert.AreEqual(XLEntityHelper.NormalizeXlsCellValue((int?)5).GetType(), typeof(int));

            Assert.AreEqual(XLEntityHelper.NormalizeXlsCellValue(5L).GetType(), typeof(long));
            Assert.AreEqual(XLEntityHelper.NormalizeXlsCellValue((long?)5L).GetType(), typeof(long));

            Assert.AreEqual(XLEntityHelper.NormalizeXlsCellValue(5D).GetType(), typeof(double));
            Assert.AreEqual(XLEntityHelper.NormalizeXlsCellValue((double?)5D).GetType(), typeof(double));

            Assert.AreEqual(XLEntityHelper.NormalizeXlsCellValue(null).GetType(), typeof(string));

            Assert.AreEqual(XLEntityHelper.NormalizeXlsCellValue(DBNull.Value).GetType(), typeof(string));
        }


        private static void WriteFileFromStream(string fileOutPut, Stream stream)
        {
            stream.Position = 0;
            byte[] buffer = new byte[stream.Length];
            stream.Read(buffer, 0, buffer.Length);

            File.WriteAllBytes(fileOutPut, buffer);
        }
    }
}

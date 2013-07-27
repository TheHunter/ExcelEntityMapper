using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using ExcelEntityMapper;
using ExcelEntityMapper.Exceptions;
using ExcelEntityMapper.Impl;
using NUnit.Framework;
using TestExcelEntityMapper.Domain;

namespace ExcelEntityMapperTest
{
    class FactoryTest
        : SheetWorkerTest
    {
        private byte[] xlsResource;
        private byte[] xlsxResource;

        public override void OnStartUp()
        {
            base.OnStartUp();
            this.xlsResource = Properties.Resources.input_XLS_;
            this.xlsxResource = Properties.Resources.input_XLSX_;
        }

        [Test]
        [Category("WrongMappers")]
        [ExpectedException(typeof(WrongParameterException))]
        public void WrongPropertymapper1()
        {
            FactoryMapper.MakeReaderPropertyMap<Person>(0, MapperType.Key, "Name", n => n.Name);
        }

        [Test]
        [Category("WrongMappers")]
        [ExpectedException(typeof(WrongParameterException))]
        public void WrongPropertymapper2()
        {
            FactoryMapper.MakeWriterPropertyMap<Person>(0, MapperType.Key, "Name", (instance, cellvalue) => instance.Name = cellvalue);
        }

        [Test]
        [Category("WrongMappers")]
        [ExpectedException(typeof(WrongParameterException))]
        public void WrongPropertymapper3()
        {
            FactoryMapper.MakeReaderPropertyMap(1, MapperType.Key, "Name", (Expression<Func<Person, string>>)null);
        }

        [Test]
        [Category("WrongMappers")]
        [ExpectedException(typeof(WrongParameterException))]
        public void WrongPropertymapper4()
        {
            FactoryMapper.MakeWriterPropertyMap(0, MapperType.Key, "Name", (Action<Person, string>)null);
        }

        [Test]
        [Category("PropertyMappers")]
        public void Propertymapper1()
        {
            var mapper = FactoryMapper.MakeReaderPropertyMap<Person>(1, MapperType.Key, "Name", n => n.Name);
            Assert.AreEqual(mapper.OperationEnabled, SourceOperation.Read);
        }

        [Test]
        [Category("PropertyMappers")]
        public void Propertymapper2()
        {
            var mapper = FactoryMapper.MakeWriterPropertyMap<Person>(1, MapperType.Key, "Name", (instance, cellvalue) => instance.Name = cellvalue);
            Assert.AreEqual(mapper.OperationEnabled, SourceOperation.Write);
        }

        [Test]
        [Category("PropertyMappers")]
        public void Propertymapper3()
        {
            var mapper = new PropertyMapper<Person>(1, MapperType.Key, "Name", (instance, cellvalue) => instance.Name = cellvalue, person => person.Name);
            Assert.AreEqual(mapper.OperationEnabled, SourceOperation.ReadWrite);
        }

        [Test]
        [Category("MakingWorkBook")]
        public void MakingWorkBookFromBytes()
        {
            var xlsWorkBook = FactoryMapper.MakeWorkBook(this.xlsResource);
            Assert.AreEqual(xlsWorkBook.Format, ExcelFormat.BIFF);

            var xlsxWorkBook = FactoryMapper.MakeWorkBook(this.xlsxResource);
            Assert.AreEqual(xlsxWorkBook.Format, ExcelFormat.XML);
        }

        [Test]
        [Category("MakingWorkBook")]
        public void MakingWorkBookFromStream()
        {
            var xlsWorkBook = FactoryMapper.MakeWorkBook(new MemoryStream(this.xlsResource));
            Assert.AreEqual(xlsWorkBook.Format, ExcelFormat.BIFF);

            var xlsxWorkBook = FactoryMapper.MakeWorkBook(new MemoryStream(this.xlsxResource));
            Assert.AreEqual(xlsxWorkBook.Format, ExcelFormat.XML);
        }
    }
}

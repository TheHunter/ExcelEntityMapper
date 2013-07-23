using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using ExcelEntityMapper;
using ExcelEntityMapper.Exceptions;
using ExcelEntityMapper.Impl;
using ExcelEntityMapper.Impl.Xml;
using NPOI.SS.UserModel;
using NUnit.Framework;
using TestExcelEntityMapper.Domain;

namespace ExcelEntityMapperTest
{
    [TestFixture]
    class PropertyMapperTest
    {
        [Test]
        [Category("WrongMappers")]
        [Description("The collection of property mappers cannot contain two mappers with identical column index.")]
        [ExpectedException(typeof(SheetParameterException))]
        public void WrongPersonalMapperTest()
        {

            List<IXLPropertyMapper<Person>> parameters = new List<IXLPropertyMapper<Person>>();

            parameters.Add(new PropertyMapper<Person>(2, MapperType.Key, (instance, cellValue) => instance.Name = cellValue, n => n.Name));
            parameters.Add(new PropertyMapper<Person>(3, MapperType.Key, (instance, cellValue) => instance.Surname = cellValue, n => n.Surname));
            parameters.Add(new PropertyMapper<Person>(4, (instance, cellValue) => instance.MarriedYear = XLEntityHelper.ToPropertyFormat<int>(cellValue), n => XLEntityHelper.ToExcelFormat(n.MarriedYear)));
            parameters.Add(new PropertyMapper<Person>(5, (instance, cellValue) => instance.BirthDate = XLEntityHelper.ToPropertyFormat<DateTime>(cellValue), n => n.BirthDate.HasValue ? n.BirthDate.Value.ToString("dd/MM/yyyy") : null));
            parameters.Add(new PropertyMapper<Person>(6, (instance, cellValue) => instance.OwnCar.Name = cellValue, n => n.OwnCar.Name));
            parameters.Add(new PropertyMapper<Person>(7, (instance, cellValue) => instance.OwnCar.Targa = cellValue, n => n.OwnCar.Targa));
            parameters.Add(new PropertyMapper<Person>(7, (instance, cellValue) => instance.OwnCar.BuildingYear = XLEntityHelper.ToPropertyFormat<int>(cellValue), n => XLEntityHelper.ToExcelFormat(n.OwnCar.BuildingYear)));

            IXLSheetMapper<Person> mapper = new XLSheetMapper<Person>(parameters);
        }

        [Test]
        [Category("WrongMappers")]
        [Description("The collection of property mapper must contain at least a key Property mapper.")]
        [ExpectedException(typeof(SheetParameterException))]
        public void WrongPersonalMapperTest2()
        {

            List<IXLPropertyMapper<Person>> parameters = new List<IXLPropertyMapper<Person>>();

            parameters.Add(new PropertyMapper<Person>(2, (instance, cellValue) => instance.Name = cellValue, n => n.Name));
            parameters.Add(new PropertyMapper<Person>(3, (instance, cellValue) => instance.Surname = cellValue, n => n.Surname));
            parameters.Add(new PropertyMapper<Person>(4, (instance, cellValue) => instance.MarriedYear = XLEntityHelper.ToPropertyFormat<int>(cellValue), n => XLEntityHelper.ToExcelFormat(n.MarriedYear)));
            parameters.Add(new PropertyMapper<Person>(5, (instance, cellValue) => instance.BirthDate = XLEntityHelper.ToPropertyFormat<DateTime>(cellValue), n => n.BirthDate.HasValue ? n.BirthDate.Value.ToString("dd/MM/yyyy") : null));
            parameters.Add(new PropertyMapper<Person>(6, (instance, cellValue) => instance.OwnCar.Name = cellValue, n => n.OwnCar.Name));
            parameters.Add(new PropertyMapper<Person>(7, (instance, cellValue) => instance.OwnCar.Targa = cellValue, n => n.OwnCar.Targa));
            parameters.Add(new PropertyMapper<Person>(8, (instance, cellValue) => instance.OwnCar.BuildingYear = XLEntityHelper.ToPropertyFormat<int>(cellValue), n => XLEntityHelper.ToExcelFormat(n.OwnCar.BuildingYear)));

            Assert.IsTrue(parameters.Count() == parameters.Count(n => n.CustomType == MapperType.Simple));

            IXLSheetMapper<Person> mapper = new XLSheetMapper<Person>(parameters);

        }

        //[Test]
        [Category("WrongMappers")]
        [ExpectedException(typeof(WrongParameterException))]
        public void WrongPropertymapper()
        {
            var mapper = new PropertyMapper<Person>(0, MapperType.Key, (instance, cellvalue) => instance.Name = cellvalue, n => n.Name);
        }


        //[Test]
        [Category("WrongMappers")]
        [ExpectedException(typeof(WrongParameterException))]
        public void WrongPropertymapper1()
        {
            //var mapper = new PropertyMapper<Person>(0, MapperType.Key, "Name", n => n.Name);
        }

        //[Test]
        [Category("WrongMappers")]
        [ExpectedException(typeof(WrongParameterException))]
        public void WrongPropertymapper2()
        {
            //var mapper = new PropertyMapper<Person>(0, MapperType.Key, "Name", (instance, cellvalue) => instance.Name = cellvalue);
        }

        //[Test]
        [Category("WrongMappers")]
        [ExpectedException(typeof(WrongParameterException))]
        public void WrongPropertymapper3()
        {
            //var mapper = new PropertyMapper<Person>(1, MapperType.Key, "Name", (Expression<Func<Person, string>>)null);
            
        }

        //[Test]
        [Category("WrongMappers")]
        [ExpectedException(typeof(WrongParameterException))]
        public void WrongPropertymapper4()
        {
            //var mapper = new PropertyMapper<Person>(0, MapperType.Key, "Name", (Action<Person, string>)null);
        }

        //[Test]
        [Category("WrongMappers")]
        public void Propertymapper1()
        {
            //var mapper = new PropertyMapper<Person>(1, MapperType.Key, "Name", n => n.Name);
            //Assert.AreEqual(mapper.OperationEnabled, SourceOperation.Read);
        }

        //[Test]
        [Category("WrongMappers")]
        public void Propertymapper2()
        {
            //var mapper = new PropertyMapper<Person>(1, MapperType.Key, "Name", (instance, cellvalue) => instance.Name = cellvalue);
            //Assert.AreEqual(mapper.OperationEnabled, SourceOperation.Write);
        }

        [Test]
        [Category("WrongMappers")]
        public void Propertymapper3()
        {
            var mapper = new PropertyMapper<Person>(1, MapperType.Key, "Name", (instance, cellvalue) => instance.Name = cellvalue, person => person.Name);
            Assert.AreEqual(mapper.OperationEnabled, SourceOperation.ReadWrite);
        }
    }
}

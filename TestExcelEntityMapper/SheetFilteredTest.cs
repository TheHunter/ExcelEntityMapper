using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ExcelEntityMapper;
using ExcelEntityMapper.Impl;
using NUnit.Framework;
using TestExcelEntityMapper.Domain;

namespace ExcelEntityMapperTest
{
    [TestFixture]
    public abstract class SheetFilteredTest
    {
        private IEnumerable<IXLPropertyMapper<Person>> propertyMappersPerson;
        private string outputPath;

        [TestFixtureSetUp]
        public virtual void OnStartUp()
        {
            outputPath = Path.GetTempPath();
            propertyMappersPerson = GetPersonMapper();
        }

        [TestFixtureTearDown]
        public void OnOverTest()
        {
            this.propertyMappersPerson = null;
            this.outputPath = null;
        }

        /// <summary>
        /// 
        /// </summary>
        public IEnumerable<IXLPropertyMapper<Person>> PropertyMappersPerson
        {
            get { return this.propertyMappersPerson; }
        }

        /// <summary>
        /// 
        /// </summary>
        public string OutputPath
        {
            get { return this.outputPath; }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="fileOutPut"></param>
        /// <param name="stream"></param>
        protected static void WriteFileFromStream(string fileOutPut, Stream stream)
        {
            stream.Position = 0;
            byte[] buffer = new byte[stream.Length];
            stream.Read(buffer, 0, buffer.Length);

            File.WriteAllBytes(fileOutPut, buffer);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public static IEnumerable<Person> GetDefaultPersons()
        {
            List<Person> persons = new List<Person>();

            Person p1 = new Person();
            p1.Name = "Mario";
            p1.Surname = "Monti";
            p1.BirthDate = DateTime.Now;
            p1.OwnCar = new Car { Name = "Alfa", Targa = "DRZ895AA", BuildingYear = 2005 };
            persons.Add(p1);

            p1 = new Person();
            p1.Name = "Silvio";
            p1.Surname = "Berlusconi";
            p1.MarriedYear = 2000;
            p1.BirthDate = DateTime.Now;
            p1.OwnCar = new Car { Name = "Audi", Targa = "DRF999BB", BuildingYear = 2009 };
            persons.Add(p1);

            return persons;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public static IEnumerable<IXLPropertyMapper<Person>> GetPersonMapper()
        {
            List<IXLPropertyMapper<Person>> parameters = new List<IXLPropertyMapper<Person>>();

            parameters.Add(new PropertyMapper<Person>(2, "Name", (instance, cellValue) => instance.Name = cellValue, n => n.Name));
            parameters.Add(new PropertyMapper<Person>(3, "Surname", (instance, cellValue) => instance.Surname = cellValue, n => n.Surname));
            parameters.Add(new PropertyMapper<Person>(4, "MarriedYear", (instance, cellValue) => instance.MarriedYear = XLEntityHelper.ToPropertyFormat<int>(cellValue), n => XLEntityHelper.ToExcelFormat(n.MarriedYear)));
            parameters.Add(new PropertyMapper<Person>(5, "DataNascita", (instance, cellValue) => instance.BirthDate = XLEntityHelper.ToPropertyFormat<DateTime>(cellValue), n => n.BirthDate.HasValue ? n.BirthDate.Value.ToString("dd/MM/yyyy") : null));
            parameters.Add(new PropertyMapper<Person>(6, (instance, cellValue) => instance.OwnCar.Name = cellValue, n => n.OwnCar.Name));
            parameters.Add(new PropertyMapper<Person>(7, (instance, cellValue) => instance.OwnCar.Targa = cellValue, n => n.OwnCar.Targa));
            parameters.Add(new PropertyMapper<Person>(8, "Anno Immatricolazione", (instance, cellValue) => instance.OwnCar.BuildingYear = XLEntityHelper.ToPropertyFormat<int>(cellValue), n => XLEntityHelper.ToExcelFormat(n.OwnCar.BuildingYear)));
            return parameters;
        }

        public static IEnumerable<IXLPropertyMapper<Person>> GetPersonMapper2()
        {
            List<IXLPropertyMapper<Person>> parameters = new List<IXLPropertyMapper<Person>>();

            parameters.Add(new PropertyMapper<Person>(2, MapperType.Key, (instance, cellValue) => instance.Name = cellValue, n => n.Name));
            parameters.Add(new PropertyMapper<Person>(3, MapperType.Key, (instance, cellValue) => instance.Surname = cellValue, n => n.Surname));
            parameters.Add(new PropertyMapper<Person>(4, (instance, cellValue) => instance.MarriedYear = XLEntityHelper.ToPropertyFormat<int>(cellValue), n => XLEntityHelper.ToExcelFormat(n.MarriedYear)));
            parameters.Add(new PropertyMapper<Person>(5, (instance, cellValue) => instance.BirthDate = XLEntityHelper.ToPropertyFormat<DateTime>(cellValue), n => n.BirthDate.HasValue ? n.BirthDate.Value.ToString("dd/MM/yyyy") : null));
            parameters.Add(new PropertyMapper<Person>(6, (instance, cellValue) => instance.OwnCar.Name = cellValue, n => n.OwnCar.Name));
            parameters.Add(new PropertyMapper<Person>(7, (instance, cellValue) => instance.OwnCar.Targa = cellValue, n => n.OwnCar.Targa));
            parameters.Add(new PropertyMapper<Person>(8, (instance, cellValue) => instance.OwnCar.BuildingYear = XLEntityHelper.ToPropertyFormat<int>(cellValue), n => XLEntityHelper.ToExcelFormat(n.OwnCar.BuildingYear)));
            return parameters;
        }
    }
}

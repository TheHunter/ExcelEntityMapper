using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelEntityMapper;
using NUnit.Framework;

namespace ExcelEntityMapperTest
{
    [TestFixture]
    class PropertyMapperTest
    {
        [Test]
        public void DefaultPersonMapperTest()
        {
            var dd = SheetFilteredTest.GetPersonMapper2();

            Assert.IsTrue(dd.Count(n => n.CustomType == MapperType.Key) == 2);
        }

    }
}

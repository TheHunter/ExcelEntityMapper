using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TestExcelEntityMapper.Domain
{
    /// <summary>
    /// 
    /// </summary>
    public class Person
    {
        public Person()
        {
            //this.OwnCar = new Car();
        }

        public string Name { get; set; }
        public string Surname { get; set; }
        public DateTime? BirthDate { get; set; }
        public Car OwnCar { get; set; }
        public int? MarriedYear {get; set;}
    }
}

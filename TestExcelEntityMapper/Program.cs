using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TestExcelEntityMapper.Domain;
using ExcelEntityMapper.Impl;
using ExcelEntityMapper;
using System.IO;
using System.Linq.Expressions;
using System.Reflection;
using System.Globalization;

namespace TestExcelEntityMapper
{
    class Program
    {
        //private static string[] DateFormats = { "dd/MM/yyyy", "dd/MM/yyyy hh:mm:ss" };
        //private static HashSet<Type> DateTypes = new HashSet<Type>();

        /// <summary>
        /// Try to execute method in order to verify each method.
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            //DateTypes.Add(typeof(DateTime));
            //DateTypes.Add(typeof(DateTime?));

            //WriteSmartExcel();
            //WriteSmartExcel_Xlsx();

            //ReadSmartExcel();
            //ReadSmartExcelFiltered();

            //WriteToCSV();

            //WriteIntoDifferentFormat();
            WriteAllSheetsToCSV();
        }

        
        static void ReadSmartExcel()
        {
            var a = TestExcelEntityMapper.Properties.Resources.output_XLS_;
            XWorkBook wb = new XWorkBook(new MemoryStream(a), ExcelFormat.Xls);
            IXLSheetWorker<Person> sheet = new XSheetMapper<Person>(1, true, GetPersonMapper());
            sheet.SheetName = "Persons";

            Dictionary<int, Person> pp = new Dictionary<int, Person>();
            int count = sheet.ReadObjects(wb, pp);

            Console.WriteLine("read rows: " + count.ToString());
            Console.WriteLine("read objects: " + pp.Count.ToString());
            Console.ReadLine();
        }

        
        static void ReadSmartExcelFiltered()
        {
            var a = TestExcelEntityMapper.Properties.Resources.output_XLS_;
            XWorkBook wb = new XWorkBook(new MemoryStream(a), ExcelFormat.Xls);

            IXLSheetFiltered<Person> sheet = new XSheetFilteredMapper<Person>(1, true, GetPersonMapper());
            sheet.SheetName = "Persons";

            Dictionary<int, Person> pp = new Dictionary<int, Person>();
            //int count = sheet.ReadFilteredObjects(wb, pp, n => n.MarriedYear != null);
            int count = sheet.ReadFilteredObjects(wb, pp, n => n.BirthDate.HasValue && n.BirthDate.Value.Year >= 1980);
            
            //int count = sheet.ReadObjects(wb, pp);
            //var lista = pp.Where(n => n.Value.Name == "Elis");

            Console.WriteLine("read rows: " + count.ToString());
            Console.WriteLine("read objects: " + pp.Count.ToString());
            Console.ReadLine();
        }

        
        static void WriteSmartExcel()
        {
            XWorkBook wb = new XWorkBook();

            IXLSheetWorker<Person> sheet = new XSheetMapper<Person>(1, true, GetPersonMapper());
            wb.AddSheet("Persons");
            IEnumerable<Person> lista = GetDefaultPersons();

            sheet.SheetName = "Persons";
            sheet.WriteObjects(wb, lista);

            Stream stream = wb.Save(ExcelFormat.Xls);
            try
            {
                WriteFileFromStream(Path.Combine(Path.GetTempPath(), "output_3.xls"), stream);
            }
            catch (Exception)
            {
                Console.WriteLine("error....");
            }
            finally
            {
                if (stream != null)
                {
                    stream.Close();
                }
            }
        }

        
        static void WriteSmartExcel_Xlsx()
        {
            XWorkBook wb = new XWorkBook();

            IXLSheetWorker<Person> sheet = new XSheetMapper<Person>(1, false, GetPersonMapper());
            wb.AddSheet("Persons");
            IEnumerable<Person> lista = GetDefaultPersons();

            sheet.SheetName = "Persons";
            sheet.WriteObjects(wb, lista);

            Stream stream = wb.Save(ExcelFormat.Xlsx);
            try
            {
                WriteFileFromStream(Path.Combine(Path.GetTempPath(), "output_3.xlsx"), stream);
            }
            catch (Exception)
            {
                Console.WriteLine("error....");
            }
            finally
            {
                if (stream != null)
                {
                    stream.Close();
                }
            }
        }


        static void WriteToCSV()
        {
            XWorkBook wb = new XWorkBook(new MemoryStream(TestExcelEntityMapper.Properties.Resources.output_XLS_), ExcelFormat.Xls);

            //Stream mem = wb.ToCSV("Persons", '\t', true, false);
            Stream mem = wb.SaveIntoCSV("Second", '\t', true, false);
            try
            {
                //wb.writeCSV(mem); //default a ','
                //wb.writeCSV(mem, '\t', true, true);
                //wb.writeCSV(mem, ';', true, true);
                //wb.writeCSV(mem, '\t', false, false);

                WriteFileFromStream(Path.Combine(Path.GetTempPath(), "output_2.csv"), mem);
            }
            catch (Exception)
            {
                Console.WriteLine("Errore...");
            }
            finally
            {
                if (mem != null)
                    mem.Close();
            }
        }


        static void WriteAllSheetsToCSV()
        {
            XWorkBook wb = new XWorkBook(new MemoryStream(TestExcelEntityMapper.Properties.Resources.output_XLS_), ExcelFormat.Xls);

            var dic = wb.SaveIntoCSV(';', true, false);
            try
            {
                if (dic != null)
                {
                    foreach(var current in dic)
                    {
                        try
                        {
                            WriteFileFromStream(Path.Combine(Path.GetTempPath(), string.Format("{0}.csv", current.Key)), current.Value);
                        }
                        catch (Exception)
                        {
                            //
                        }
                        finally
                        {
                            current.Value.Close();
                        }
                    }
                }
            }
            catch (Exception)
            {
                Console.WriteLine("Errore...");
            }
            
        }


        static void WriteIntoDifferentFormat()
        {
            XWorkBook wb_XLS = new XWorkBook(new MemoryStream(TestExcelEntityMapper.Properties.Resources.output_XLS_), ExcelFormat.Xls);
            XWorkBook wb_XLSX = new XWorkBook(new MemoryStream(TestExcelEntityMapper.Properties.Resources.output_XLSX_), ExcelFormat.Xlsx);

            Stream m1 = null;
            Stream m2 = null;

            try
            {
                m1 = wb_XLS.Save(ExcelFormat.Xlsx);
                m2 = wb_XLS.Save(ExcelFormat.Xls);

                WriteFileFromStream(Path.Combine(Path.GetTempPath(), "output_from_XLS.xlsx"), m1);
                WriteFileFromStream(Path.Combine(Path.GetTempPath(), "output_from_XLSX.xls"), m2);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (m1 != null)
                {
                    m1.Close();
                }

                if (m2 != null)
                {
                    m2.Close();
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="fileOutPut"></param>
        /// <param name="stream"></param>
        static void WriteFileFromStream(string fileOutPut, Stream stream)
        {
            stream.Position = 0;
            byte[] buffer = new byte[stream.Length];
            stream.Read(buffer, 0, buffer.Length);

            File.WriteAllBytes(fileOutPut, buffer);
        }


        static IEnumerable<IXLPropertyMapper<Person>> GetPersonMapper()
        {
            List<IXLPropertyMapper<Person>> parameters = new List<IXLPropertyMapper<Person>>();
            parameters.Add(new XLPropertyMapper<Person>(1, "Name", (n, r) => n.Name = r, n => n.Name));
            parameters.Add(new XLPropertyMapper<Person>(2, "Surname", (n, r) => n.Surname = r, n => n.Surname));
            parameters.Add(new XLPropertyMapper<Person>(3, "MarriedYear", (n, r) => n.MarriedYear = XLEntityHelper.ToPropertyFormat<int>(r), n => XLEntityHelper.ToExcelFormat(n.MarriedYear)));
            parameters.Add(new XLPropertyMapper<Person>(4, "DataNascita", (n, r) => n.BirthDate = XLEntityHelper.ToPropertyFormat<DateTime>(r), n => n.BirthDate.HasValue ? n.BirthDate.Value.ToString("dd/MM/yyyy") : null));
            parameters.Add(new XLPropertyMapper<Person>(5, (n, r) => n.OwnCar.Name = r, n => n.OwnCar.Name ));
            parameters.Add(new XLPropertyMapper<Person>(6, (n, r) => n.OwnCar.Targa = r, n => n.OwnCar.Targa));
            parameters.Add(new XLPropertyMapper<Person>(7, "Anno Immatricolazione", (n, r) => n.OwnCar.BuildingYear = XLEntityHelper.ToPropertyFormat<int>(r), n => XLEntityHelper.ToExcelFormat(n.OwnCar.BuildingYear)));
            return parameters;
        }


        static IEnumerable<Person> GetDefaultPersons()
        {
            List<Person> persons = new List<Person>();

            Person p1 = new Person();
            p1.Name = "Mario";
            p1.Surname = "Monti";
            p1.BirthDate = DateTime.Now;
            p1.OwnCar = new Car { Name = "Alfa", Targa="DRZ895AA", BuildingYear=2005 };
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


    }
}

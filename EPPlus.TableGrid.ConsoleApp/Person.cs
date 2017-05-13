using System;

namespace EPPlus.TableGrid.ConsoleApp
{
    public class Person
    {
        public Person(string firstName, string address, DateTime birthdate, decimal budget)
        {
            FirstName = firstName;
            Address = address;
            Birthdate = birthdate;
            Budget = budget;
        }

        public string FirstName { get; set; }
        public string Address { get; set; }
        public DateTime Birthdate { get; set; }
        public decimal Budget { get; set; }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Entity
{
    public class Person
    {
        int id; string name; DateTime birthDate;

        public int Id { get => id; set => id = value; }
        public string Name { get => name; set => name = value; }
        public DateTime BirthDate { get => birthDate; set => birthDate = value; }

        public Person(int id, string name, DateTime birthDate)
        {
            this.id = id;
            this.name = name;
            this.birthDate = birthDate;
        }
    }
}

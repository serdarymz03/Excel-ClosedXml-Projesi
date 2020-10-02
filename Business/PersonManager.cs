using DataAccess;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Business
{
    public class PersonManager
    {
        static PersonManager personManager;
        PersonDal personDal;

        public PersonManager()
        {
            personDal = PersonDal.GetInstance();
        }

        //Get People As DataTable

        public DataTable GetPeopleAsTable()
        {
            try
            {
                return personDal.GetPeopleAsTable();
            }
            catch
            {
                return null;
            }
        }



        public static PersonManager GetInstance()
        {
            if (personManager == null)
            {
                personManager = new PersonManager();
            }
            return personManager;
        }
    }
}

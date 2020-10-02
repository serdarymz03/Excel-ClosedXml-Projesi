using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataAccess
{
    public class PersonDal
    {
        static PersonDal personDal;
        SqlService sqlService;
        private PersonDal()
        {
            sqlService = SqlService.GetInstance();
        }

        //Get People As DataTable
        public DataTable GetPeopleAsTable()
        {
            try
            {
                return sqlService.GetDataTable("PeopleList");
            }
            catch
            {
                return null;
            }
        }



        public static PersonDal GetInstance()
        {
            if (personDal==null)
            {
                personDal = new PersonDal();
            }
            return personDal;
        }
    }
}

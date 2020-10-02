using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataAccess
{
    public class SqlService
    {
        static SqlService sqlService;
        SqlConnection connection;
        string connectionString = ConfigurationManager.ConnectionStrings["sqlConnect"].ConnectionString;

        private SqlService()
        {
            connection = new SqlConnection();
            connection.ConnectionString = connectionString;
        }
        SqlConnection OpenConnection()
        {
            if (connection.State == ConnectionState.Closed)
            {
                connection.Open();
            }
            return connection;
        }

        void CloseConnection()
        {
            if (connection.State == ConnectionState.Open)
            {
                connection.Close();
            }
        }

        //Get DataTable Method

        public DataTable GetDataTable(string commandText, params SqlParameter[] parameters)
        {
            SqlCommand command = new SqlCommand();
            command.CommandType = CommandType.StoredProcedure;
            command.CommandText = commandText;
            command.Connection = connection;

            if (parameters != null)
            {
                command.Parameters.AddRange(parameters);
            }

            SqlDataAdapter dataAdapter = new SqlDataAdapter();
            dataAdapter.SelectCommand = command;
            DataTable dataTable = new DataTable();
            dataAdapter.Fill(dataTable);
            return dataTable;
        }


        public static SqlService GetInstance()
        {
            if (sqlService == null)
            {
                sqlService = new SqlService();
            }
            return sqlService;
        }
    }
}

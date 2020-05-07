using Dapper;
using MySql.Data.MySqlClient;
using System.Collections.Generic;
using System.Data;

namespace ExcelDnaSampleQuerySQLServer
{
    public static class DBHelper
    {      
        

        private static IDbConnection GetConnection(string connectionString)
        {
            return new MySqlConnection(connectionString);
        }


        public static IEnumerable<dynamic> QueryTable(string query,string connectionString)
        {
            using (var connection = GetConnection(connectionString))
            {
                var result = connection.Query(query);
                return result;
            }
        }

    }
}

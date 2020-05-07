using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using System;
using System.Runtime.Caching;

namespace ExcelDnaSampleQuerySQLServer
{    

    public static class MyFunctions
    {
        // this is a quick demo to show how to query dynamically with ExcelDNA & Dapper
        // I'm extracting the column names from the DapperRow, therfore the cast to IDictionary<string,object>.https://dapper-tutorial.net/knowledge-base/26932789/dapper-results-dapper-row--with-bracket-notation
        // Don't use this in production, SQL injections etc.
      

        [ExcelFunction(Description = "Generic Function", Category = "My Functions")]
        public static object Query(
            [ExcelArgument(Name ="SQL Query", Description ="SQL Query")] string query,
            [ExcelArgument(Name = "SQL Connection String", Description = "SQL Connection String")] string connectionString,
            [ExcelArgument(Name = "optional caching in secs", Description = "enter caching time in seconds, default is 5")] int caching

            )
        {
            int cachingDurationInSeconds = caching == 0 ? 5 : caching;

            try
            {
                var data = QueryWithCaching(query, connectionString, cachingDurationInSeconds);
                var result = data.ToArray();

                var rowsCount = result.Length;
                var cols = result[0].Keys.ToArray();
                var colsCount = result[0].Keys.Count;
                var a = new object[rowsCount + 1, colsCount];

                for (int i = 0; i < colsCount; i++)
                {
                    a[0, i] = cols[i];
                }

                for (int i = 0; i < rowsCount; i++)
                {
                    for (int y = 0; y < colsCount; y++)
                    {
                        var colKey = cols[y];
                        a[i + 1, y] = result[i][colKey];
                    }
                }

                return ArrayResizer.Resize(a);
            }
            catch (Exception ex)
            {

                return ex.Message;
            }            
          
        }    
        

        private static IEnumerable<IDictionary<string, object>> QueryWithCaching(string query,string connectionString,int cachingDuractionInSeconds)
        {
            ObjectCache cache = MemoryCache.Default;

           

            if (!(cache[query] is IEnumerable<IDictionary<string, object>> repsonse))
            {
                var cachePolicy = new CacheItemPolicy();                

                cachePolicy.AbsoluteExpiration = DateTimeOffset.Now.AddSeconds(cachingDuractionInSeconds);

                repsonse = DBHelper.QueryTable(query, connectionString) as IEnumerable<IDictionary<string, object>>;


                cache.Set(query, repsonse, cachePolicy);

            }               


            return repsonse;

        }
       
    }       
}

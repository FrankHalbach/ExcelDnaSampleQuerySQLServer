using ExcelDna.Integration;
using ExcelDna.IntelliSense;

namespace ExcelDnaSampleQuerySQLServer
{
    public class AddInLoad : IExcelAddIn
    {
        public static object XLApp { get; internal set; }

        public void AutoOpen()
        {
            // loading intelli server in debug gives deadlock
#if !DEBUG
            IntelliSenseServer.Install();
#endif




        }


        public void AutoClose()
        {

#if !DEBUG
            IntelliSenseServer.Uninstall();           
#endif

        }
    }
}

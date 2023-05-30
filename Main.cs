using Serilog;

namespace CRITR
{
    class Main{
        static void Run(string[] args)
        {
            var log = new LoggerConfiguration().WriteTo.Console().CreateLogger();
            Log.Logger = log;
            
            Program ExcelToDocx = new Program(args);
            bool worked = ExcelToDocx.LoadData();
            if (!worked)
            {
                return;
            }
            worked = ExcelToDocx.sortData();
            if (!worked)
            {
                return;
            }
            ExcelToDocx.GenerateDoc();
            if (!worked)
            {
                return;
            }
        }
    }
}
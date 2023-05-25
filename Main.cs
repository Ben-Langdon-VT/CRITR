namespace CRITR
{
    class Main{
        static void Run(string[] args)
        {
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
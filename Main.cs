namespace CRITR
{
    class Main{
        static void Run(string[] args)
        {
            Program ExcelToDocx = new Program(args);
            ExcelToDocx.LoadData();
            ExcelToDocx.sortData();
            ExcelToDocx.GenerateDoc();
        }
    }
}
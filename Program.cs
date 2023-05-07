using System;

namespace CRITR{
    class Program
    {
        
        static void Main(string[] args)
        {
            string filePath = @"C:\Users\Ben\Desktop\test\testdoc.xlsx";
            ExcelFileHandler excelHandler = new ExcelFileHandler(filePath);
            Console.Write(excelHandler.PrintExcelFile());
        }
    }
}
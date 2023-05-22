using System;

namespace CRITR{
    class ExcelFileHandlerTest
    {
        public ExcelFileHandlerTest()
        {
            
        }
        public void Main()
        {
            string filePath = @"C:\Users\Ben\Desktop\test\testdoc.xlsx";
            ExcelFileHandler excelHandler = new ExcelFileHandler(filePath);
            Console.Write(excelHandler.PrintExcelFile());
            Console.WriteLine();
            Console.Write(excelHandler.PrintRows());
        }
    }
}
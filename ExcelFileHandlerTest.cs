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
            ExcelFileHandler excelHandler = new ExcelFileHandler();
            excelHandler.Open(filePath);
            excelHandler.Dispose();
        }
    }
}
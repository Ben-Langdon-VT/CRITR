using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;


//reference link for setting up .net/openxml from nuget/dotnet
//https://www.nuget.org/packages/DocumentFormat.OpenXml/#supportedframeworks-body-tab
namespace CRITR
{
    class ExcelFileHandler
    {
        public string fileName = "default.xlsx";
        SpreadsheetDocument excelDoc;

        public ExcelFileHandler(string _fileName)
        {
            fileName = _fileName;
            excelDoc = SpreadsheetDocument.Open(fileName, false);

        }

        public string PrintExcelFile()
        {
            string output = "";
            WorkbookPart? workbookPart = excelDoc.WorkbookPart;

            if (workbookPart == null)
            {
                return "no workbookPart found in file " + fileName;
            }

            Sheet? sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault();

            if (sheet == null)
            {
                return "no sheets found in file " + fileName;
            }

            WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

            foreach (Row row in sheetData.Elements<Row>())
            {
                
                foreach (Cell cell in row.Elements<Cell>())
                {
                    output += cell.InnerText + " ";
                    var value = cell.CellValue;
                    if (value == null){
                        Console.WriteLine("Null");
                        continue;
                    }
                    Console.WriteLine(cell.CellValue.InnerText);
                }
                output += "\n";
            }
            return output;
        }

    }
}
using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;


//reference link for setting up .net/openxml from nuget/dotnet
//https://www.nuget.org/packages/DocumentFormat.OpenXml/#supportedframeworks-body-tab
namespace CRITR
{

    //Assumptions: simple workbook, no advanced parts, has only 1 worksheet
    class ExcelFileHandler
    {
        //internal stuff so i only need to lookup file once
        public string fileName = "default.xlsx";
        SpreadsheetDocument excelDoc;
        SharedStringTablePart sstPart;
        SharedStringTable ssTable;
        WorksheetPart worksheetPart;
        SheetData sheetData;

        public ExcelFileHandler(string _fileName)
        {
            fileName = _fileName;
            excelDoc = SpreadsheetDocument.Open(fileName, false);
            WorkbookPart? workbookPart = excelDoc.WorkbookPart;
            //several nullable types, throw an error if one of them is null. will probably only happen if pointed at a weird/damaged/not xlsx file or Oxml standard changes
            if (workbookPart == null)
            {
                throw new ArgumentException("no workbookPart found in file " + fileName);
            }
            //We assume that we are only using worsheet 1
            Sheet? sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault();

            if (sheet == null)
            {
                throw new ArgumentException("no sheets found in file " + fileName);
            }

            sstPart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
            ssTable = sstPart.SharedStringTable;

            StringValue? id = sheet.Id;

            if (id == null)
            {
                throw new AggregateException("something wrong with file format, could not find sheet id in " + fileName);
            }

            worksheetPart = (WorksheetPart)workbookPart.GetPartById(id);
            sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

        }
        //other print option, somewhat lame, riddled with possible nullable types so probs dont use much/rework if i need more excel debugging.
        public string PrintExcelFile()
        {
            string output = "";


            foreach (Row row in sheetData.Elements<Row>())
            {
                output += row.RowIndex + ": ";

                foreach (Cell cell in row.Elements<Cell>())
                {
                    string cellname = cell.CellReference;
                    Console.WriteLine("innerText: {0}, DataType: {1}", cell.InnerText, cell.DataType);
                    if (cell.DataType == null)
                    {
                        output += cellname + ": " + cell.InnerText + ", ";

                    }
                    else if (cell.DataType != CellValues.SharedString)
                    {
                        output += cell.InnerText + ", ";
                    }
                    else if (cell.DataType == CellValues.SharedString)
                    {
                        int ssid = int.Parse(cell.CellValue.Text);
                        string str = ssTable.ChildElements[ssid].InnerText;
                        output += cellname + ": " + str + ", ";
                    }
                }
                output += "\n";
            }
            return output;
        }
        //simple print rows functions
        public string PrintRows()
        {
            string outstring = "";
            var rows = sheetData.Descendants<Row>();
            foreach (Row row in rows)
            {
                outstring += row.RowIndex + ", ";
            }
            return outstring;
        }

        //stolen from stack overflow, i wrote a version that was recursive and indexes a string of values which is bad for memory/reliability,
        //this one uses a loop(no accidental stack overflow) and directly generates char value so its a bit better stability/memory wise.
        private string columnNumToAlpha(int columnNumber)
        {
            int dividend = (int)columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }
        //Generate cell name from int row/int column, throws error if below 1 value for either
        private string CellName(int row, int column)
        {
            if (row <= 0 || column <= 0)
            {
                throw new ArgumentException("Bad row column in cellname converter, smallest valid row and column value is 1");
            }
            string columnName = columnNumToAlpha(column);
            return columnName + row.ToString();
        }

        //These are all nullable, couldnt think of a better way of handling it
        //i really want to return an empty cell/row here with the right name/ datatype /value but cant declare my own cells

        public Cell? lookupCellNumber(int row, int column)
        {
            //Again assumes it is using worksheet1            
            string cellName = CellName(row, column);
            Cell? theCell = worksheetPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == cellName).FirstOrDefault();

            return theCell;
        }


        //searches for cell in given row, slightly faster for big docs than searching for cell in whole thing
        public Cell? lookupCellNumberRow(int rowNum, int columnNum, Row row)
        {

            string cellName = CellName(rowNum, columnNum);
            Cell? theCell = row.Descendants<Cell>().Where(c => c.CellReference == cellName).FirstOrDefault();
            return theCell;
        }

        //search for row in worksheet1
        public Row? lookupRow(int rowNum)
        {
            Row? row = worksheetPart.Worksheet.Descendants<Row>().Where(r => r.InnerText == rowNum.ToString()).FirstOrDefault();
            return row;
        }

        public String GetCellValueAsString(Cell cell)
        {
            String outstring = "";
            if (cell.DataType == null)
            {
                outstring = cell.InnerText;

            }
            else if (cell.DataType != CellValues.SharedString)
            {
                outstring = cell.InnerText;
            }
            else if (cell.DataType == CellValues.SharedString)
            {
                int ssid = int.Parse(cell.CellValue.Text);
                outstring = ssTable.ChildElements[ssid].InnerText;
            }

            return outstring;
        }
        //Get Headings row
        public List<String> GetHeaders(int headings)
        {
            List<String> headers = new List<String>();
            Row? headerRow = lookupRow(1);
            if (headerRow == null)
            {
                throw new ArgumentNullException("Row 1 (header row) is empty (null)");
            }
            int rowCount = headerRow.Descendants<Cell>().Count<Cell>();
            if (rowCount != headings) throw new FormatException(String.Format("Supposed to be {0} cells in this row but {1} detected", headings, rowCount));

            for (int i = 1; i <= headings; i++)
            {
                Cell? c = lookupCellNumberRow(headings, i, headerRow);
                if (c == null) throw new ArgumentNullException(String.Format("Empty cell detected at {0}", CellName(headings, i)));
                headers.Add(GetCellValueAsString(c));
            }

            return headers;
        }
        //Get count of rows so we can make a loop in 
        public int CountDataRows()
        {
            return sheetData.Descendants<Row>().Where(r => r.InnerText != "1").Count<Row>();
        }

        //Read in a data Row
        public List<String> GetDataRow(int headings, int rowNum)
        {
            if (rowNum < 0) throw new ArgumentException("data row number must be greater than 0");
            rowNum += 2;
            List<String> data = new List<String>();
            Row? dataRow = lookupRow(rowNum);
            if (dataRow == null)
            {
                throw new ArgumentNullException(String.Format("Row {0} is empty (null)"), rowNum.ToString());
            }
            int rowCount = dataRow.Descendants<Cell>().Count<Cell>();
            if (rowCount != headings) throw new FormatException(String.Format("Supposed to be {0} cells in this row but {1} detected", headings, rowCount));

            for (int i = 1; i <= headings; i++)
            {
                Cell? c = lookupCellNumberRow(headings, i, dataRow);
                if (c == null) throw new ArgumentNullException(String.Format("Empty cell detected at {0}", CellName(headings, i)));
                data.Add(GetCellValueAsString(c));
            }

            return data;

        }
    }
}
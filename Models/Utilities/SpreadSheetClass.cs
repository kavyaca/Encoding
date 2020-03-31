using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;



namespace CSV.Models.Utilities
{
    public class SpreadSheetClass
    {



        public static void CreateSpreadsheetWorkbook(string filepath, List<Student> st)
        {
            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.
                Create(filepath, SpreadsheetDocumentType.Workbook);

            // Add a WorkbookPart to the document.
            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
                AppendChild<Sheets>(new Sheets());


            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.
                GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "Sheet1"


            };




            sheets.Append(sheet);





            AddUpdateCellValue(spreadsheetDocument, "Sheet1", 1, "A", "Hello, my name is Kavya", "str");



            workbookpart.Workbook.Save();


            spreadsheetDocument.Close();


            InsertWorksheet(filepath, st);

        }







        private static void AddUpdateCellValue(SpreadsheetDocument spreadSheet, string sheetname,
    uint rowIndex, string columnName, string text, string value)
        {
            // Opening document for editing            
            WorksheetPart worksheetPart =
             RetrieveSheetPartByName(spreadSheet, sheetname);

            if (worksheetPart != null)
            {
                Cell cell = InsertCellInSheet(columnName, (rowIndex + 1), worksheetPart);
                cell.CellValue = new CellValue(text);

                //cell datatype
                if (value.Equals("str"))
                {
                    cell.DataType =
                     new EnumValue<CellValues>(CellValues.String);
                }

                if (value.Equals("num"))
                {
                    cell.DataType =
                new EnumValue<CellValues>(CellValues.Number);
                }

                if (value.Equals("dat"))
                {
                    cell.DataType =
          new EnumValue<CellValues>(CellValues.Date);
                }

                if (value.Equals("b"))
                {
                    cell.DataType =
                new EnumValue<CellValues>(CellValues.Boolean);
                }


               


                // Save the worksheet.            
                worksheetPart.Worksheet.Save();
            }
        }

        public static void InsertWorksheet(string docName, List<Student> st)
        {
            // Open the document for editing.
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true))
            {
                // Add a blank WorksheetPart.
                WorksheetPart newWorksheetPart = spreadSheet.WorkbookPart.AddNewPart<WorksheetPart>();
                newWorksheetPart.Worksheet = new Worksheet(new SheetData());

                Sheets sheets = spreadSheet.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                string relationshipId = spreadSheet.WorkbookPart.GetIdOfPart(newWorksheetPart);

                // Get a unique ID for the new worksheet.
                uint sheetId = 1;
                if (sheets.Elements<Sheet>().Count() > 0)
                {
                    sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                }

                // Give the new worksheet a name.
                string sheetName = "Sheet" + sheetId;

                // Append the new worksheet and associate it with the workbook.
                Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
                sheets.Append(sheet);


            }
            AddHeaderDataInExcel(docName);
            AddDataInExcel(docName, st);
        }


        private static WorksheetPart RetrieveSheetPartByName(SpreadsheetDocument document,
         string sheetName)

        {

            IEnumerable<Sheet> sheets =
             document.WorkbookPart.Workbook.GetFirstChild<Sheets>().
            Elements<Sheet>().Where(s => s.Name == sheetName);
            if (sheets.Count() == 0)
                return null;

            string relationshipId = sheets.First().Id.Value;
            WorksheetPart worksheetPart = (WorksheetPart)
            document.WorkbookPart.GetPartById(relationshipId);
            Console.WriteLine("hi test:: " + worksheetPart);
            return worksheetPart;
        }

        //insert cell in sheet based on column and row index            
        private static Cell InsertCellInSheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;
            Row row;
            //check whether row exist or not            
            //if row exist            
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            //if row does not exist then it will create new row            
            else
            {
                row = new Row()
                {
                    RowIndex = rowIndex
                };
                sheetData.Append(row);
            }
            //check whether cell exist or not            
            //if cell exist            
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            //if cell does not exist            
            else
            {
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }
                Cell newCell = new Cell()
                {
                    CellReference = cellReference
                };
                row.InsertBefore(newCell, refCell);
                worksheet.Save();
                return newCell;
            }
        }


        private static void AddHeaderDataInExcel(string docName)
        {


            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true))
            {

                AddUpdateCellValue(spreadSheet, "Sheet2", 0, "A", "UniqueId","str");
                AddUpdateCellValue(spreadSheet, "Sheet2", 0, "B", "StudentId", "num");
                AddUpdateCellValue(spreadSheet, "Sheet2", 0, "C", "FirstName", "str");
                AddUpdateCellValue(spreadSheet, "Sheet2", 0, "D", "LastName", "str");
                AddUpdateCellValue(spreadSheet, "Sheet2", 0, "E", "DateOfBirth", "dat");
                AddUpdateCellValue(spreadSheet, "Sheet2", 0, "F", "IsMe", "num");
                AddUpdateCellValue(spreadSheet, "Sheet2", 0, "G", "Age", "num");


            }

        }

            private static void AddDataInExcel(string docName, List<Student> st)
        {

           
            int value = 0;
            foreach (var data in st)
            {
                value++;
                uint res = Convert.ToUInt32(value);

                using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true))
                {


                   int a = DateTime.Today.Year - data.DateOfBirthDT.Year;
                    if (a > 100)
                    {
                        a = 24;
                    }

                    String isMe = "0";
                 
                   

                    if(data.MyRecord == true)
                    {
                        isMe = "1";
                    }

                    Console.WriteLine("ISME>>>> " + isMe);
                    AddUpdateCellValue(spreadSheet, "Sheet2", res, "A", (Guid.NewGuid()).ToString(), "str");
                    AddUpdateCellValue(spreadSheet, "Sheet2", res, "B", data.StudentId, "num");
                    AddUpdateCellValue(spreadSheet, "Sheet2", res, "C", data.FirstName, "str");
                    AddUpdateCellValue(spreadSheet, "Sheet2", res, "D", data.LastName, "str");
                    AddUpdateCellValue(spreadSheet, "Sheet2", res, "E", data.DateOfBirth, "dat");
                    AddUpdateCellValue(spreadSheet, "Sheet2", res, "F", isMe, "num");
                    AddUpdateCellValue(spreadSheet, "Sheet2", res, "G",a.ToString(), "num");




                }

            }
        }
    }
}
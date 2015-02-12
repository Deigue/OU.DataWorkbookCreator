using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OU.DataWorkbookCreator
{
    class SpreadsheetHelper
    {
        public WorksheetPart GetWorkSheetPart(WorkbookPart workbookPart, string sheetName)
        {
            string relevantId = workbookPart.Workbook.Descendants<Sheet>().First(s => sheetName.Equals(s.Name)).Id;
            return (WorksheetPart)workbookPart.GetPartById(relevantId);
        }

        private Row GetRow(WorksheetPart worksheetPart, uint correction, uint rowIndex)
        {
            if (rowIndex <= worksheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>().Count() + correction)
            {
                return worksheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>().Where(x => x.RowIndex == rowIndex).First();
            }
            else
                return null;
        }

        private Cell GetCell(WorksheetPart worksheetPart, uint correction, uint rowIndex, char column)
        {
            Row row = GetRow(worksheetPart, correction, rowIndex);

            //New Row
            if (row == null)
            {                
                row = new Row() { RowIndex = rowIndex };
                Cell cell = new Cell();
                string cellIndex = column + rowIndex.ToString();
                cell.CellReference = cellIndex;
                row.AppendChild(cell);
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                sheetData.AppendChild(row);
                return cell;
            }
            //Cell exists in Row
            else if((int)column-64 <= row.Elements<Cell>().Count())
            {
                return row.Elements<Cell>().Where(cell => string.Compare(cell.CellReference.Value, column.ToString() + rowIndex, true)==0).First();
            }
            //Inserts Cell into existing Row
            else
            {
                Cell cell = new Cell();
                string cellIndex = column + rowIndex.ToString();
                cell.CellReference= cellIndex;
                row.AppendChild(cell);
                return cell;
            }
        }

        public void UpdateCell(WorksheetPart worksheetPart, uint correction, string cellData, uint rowIndex, char column, int cellType, uint style)
        {
            Cell cell = GetCell(worksheetPart, correction, rowIndex, column);
            cell.CellValue = new CellValue(cellData);
            cell.StyleIndex = style;

            if (cellType == 1)
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
            if (cellType == 2)
                cell.DataType = new EnumValue<CellValues>(CellValues.String);
        }

        public void UpdateCellFormula(WorksheetPart worksheetPart, uint correction, uint rowIndex, char column, int cellType, uint style, string formula)
        {
            Cell cell = GetCell(worksheetPart, correction, rowIndex, column);

            if (cellType == 1)
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
            if (cellType == 2)
                cell.DataType = new EnumValue<CellValues>(CellValues.String);

            CellFormula cellFormula = new CellFormula(formula);
            cell.StyleIndex = style;
            cellFormula.CalculateCell = true;
            CellValue value = new CellValue();
            value.Text = "0";
            cell.Append(cellFormula);
            cell.Append(value);
        }
    }//Class
}//Namespace

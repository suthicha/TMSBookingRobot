using ClosedXML.Excel;
using System;
using System.IO;

namespace TMSBookingRobot.Controllers
{
    public class ExcelController
    {
        private IXLWorksheet _worksheet;
        private XLWorkbook _workbook;
        private int _lastRow;

        public ExcelController()
        {
        }

        public void CreateWorkSheet(string sheetName)
        {
            if (_workbook == null)
                _workbook = new XLWorkbook();

            _worksheet = _workbook.Worksheets.Add(sheetName);
            _lastRow = 0;
        }

        public void CreateHeader(params string[] title)
        {
            for (int i = 0; i < title.Length; i++)
            {
                var colIndex = i + 1;
                _worksheet.Cell(1, colIndex).Value = title[i].ToUpper();
                _worksheet.Cell(1, colIndex).Style.Font.Bold = true;
                _worksheet.Cell(1, colIndex).Style.Border.BottomBorder = XLBorderStyleValues.Double;
            }
            _lastRow++;
        }

        public void AddRow(params Object[] data)
        {
            for (int i = 0; i < data.Length; i++)
            {
                dynamic obj = data[i];
                var colIndex = i + 1;
                var cellstyle = (XLCellValues)Enum.Parse(typeof(XLCellValues), obj.CellType);
                _worksheet.Cell(_lastRow + 1, colIndex).Value = obj.Tag;

                if (cellstyle == XLCellValues.Text)
                {
                    _worksheet.Cell(_lastRow + 1, colIndex).DataType = XLCellValues.Text;
                }
                else
                {
                    _worksheet.Cell(_lastRow + 1, colIndex).DataType = XLCellValues.DateTime;
                    _worksheet.Cell(_lastRow + 1, colIndex).Style.DateFormat.Format = "dd-MMM-yyyy HH:mm";
                }
            }
            _lastRow++;
        }

        private void DrawLine(int col)
        {
            for (int i = 0; i < col; i++)
            {
                var colIndex = i + 1;
                _worksheet.Cell(_lastRow, colIndex).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            }
        }

        public void AdjustWorksheetContent(params int[] widths)
        {
            _worksheet.Rows().AdjustToContents();
            _worksheet.Columns().AdjustToContents();

            for (int i = 0; i < widths.Length; i++)
            {
                _worksheet.Column(i + 1).Width = widths[i];
            }
        }

        public void Save(string outputPath)
        {
            if (File.Exists(outputPath))
                File.Delete(outputPath);

            _worksheet.Rows().AdjustToContents();
            _worksheet.Columns().AdjustToContents();
            _workbook.SaveAs(outputPath);
        }
    }
}
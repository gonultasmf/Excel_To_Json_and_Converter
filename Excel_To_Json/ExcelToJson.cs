using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDataReader;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Data;
using System.Text;

namespace Excel_To_Json;

public static class ExcelToJson
{
    public static void ToJson(string filePath, string outputFileName)
    {
        using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filePath, false))
        {
            WorkbookPart workbookPart = doc.WorkbookPart;
            var sb = new StringBuilder();
            sb.AppendLine("{");
            
            foreach (Sheet sheet in workbookPart.Workbook.Descendants<Sheet>())
            {
                WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                Worksheet worksheet = worksheetPart.Worksheet;
                SheetData sheetData = worksheet.GetFirstChild<SheetData>();

                List<Dictionary<string, string>> data = new List<Dictionary<string, string>>();

                Row headerRow = sheetData.Descendants<Row>().FirstOrDefault();
                List<string> columnNames = new List<string>();
                foreach (Cell cell in headerRow.Descendants<Cell>())
                {
                    columnNames.Add(GetCellValue(workbookPart, cell));
                }

                foreach (Row row in sheetData.Descendants<Row>())
                {
                    Dictionary<string, string> rowData = new Dictionary<string, string>();
                    int columnIndex = 0;
                    foreach (Cell cell in row.Descendants<Cell>())
                    {
                        string columnName = columnNames[columnIndex];
                        string cellValue = GetCellValue(workbookPart, cell);
                        rowData.Add(columnName, cellValue);
                        columnIndex++;
                    }
                    data.Add(rowData);
                }

                string sheetName = sheet.Name;
                string jsonVal = JsonConvert.SerializeObject(data, Formatting.Indented);
                sb.AppendLine("\"" + sheetName + "\": " + jsonVal + ",");
            }
            sb.AppendLine("}");
            JObject dataJson = JObject.Parse(sb.ToString());
            JObject updatedData = new JObject();

            foreach (var month in dataJson)
            {
                var monthData = (JArray)month.Value;
                JArray updatedMonthData = new JArray();

                for (int i = 0; i < monthData.Count; i++)
                {
                    var rowLabels = monthData[i]["Row Labels"].ToString();
                    if (IsSixUppercaseLetters(rowLabels))
                    {
                        var subRows = new JArray();
                        for (int j = i + 1; j < monthData.Count; j++)
                        {
                            var subRowLabels = monthData[j]["Row Labels"].ToString();
                            if (IsNumeric(subRowLabels))
                            {
                                subRows.Add(monthData[j]);
                            }
                            else
                            {
                                break;
                            }
                        }
                        monthData[i]["SubRows"] = subRows;
                        updatedMonthData.Add(monthData[i]);
                    }
                }

                updatedData[month.Key] = updatedMonthData;
            }

            File.WriteAllText($"{outputFileName}.json", updatedData.ToString());
        }


    }

    private static bool IsSixUppercaseLetters(string value)
    {
        if (value.Length != 6)
            return false;

        foreach (char c in value)
        {
            if (!char.IsUpper(c))
                return false;
        }

        return true;
    }

    private static bool IsNumeric(string value)
    {
        return int.TryParse(value, out _);
    }

    static string GetCellValue(WorkbookPart workbookPart, Cell cell)
    {
        SharedStringTablePart stringTablePart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
        string value = cell.InnerText;
        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
        {
            if (stringTablePart != null)
            {
                value = stringTablePart.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
            }
        }
        return value;
    }

    private static string GenerateJson(params IXLWorksheet[] worksheets)
    {
        var jsonData = new Dictionary<string, List<Dictionary<string, string>>>();
        foreach (var worksheet in worksheets)
        {
            var worksheetName = worksheet.Name;
            var rows = worksheet.RowsUsed();

            if (rows.Count() > 1)
            {
                var columnNames = rows.First().Cells().Select(c => c.Value.ToString()).ToList();
                var rowDataList = new List<Dictionary<string, string>>();

                foreach (var row in rows.Skip(1))
                {
                    var rowData = new Dictionary<string, string>();
                    var cells = row.Cells();

                    for (int i = 0; i < columnNames.Count; i++)
                    {
                        rowData.Add(columnNames[i], cells.ElementAt(i).Value.ToString());
                    }

                    rowDataList.Add(rowData);
                }

                jsonData.Add(worksheetName, rowDataList);
            }
        }

        return JsonConvert.SerializeObject(jsonData, Formatting.Indented);
    }
}
//rowDataList.Select(row => row.ToDictionary(kvp => kvp.Key, kvp => kvp.Value.ToString())
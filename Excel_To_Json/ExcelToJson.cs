using ClosedXML.Excel;
using Newtonsoft.Json;

namespace Excel_To_Json;

public static class ExcelToJson
{
    public static void ToJson(string filePath, string outputFileName, string worksheetName = null)
    {
        using (var workbook = new XLWorkbook(filePath)) // Excel dosyası açıldı
        {
            string json = string.Empty;

            if (!string.IsNullOrEmpty(worksheetName))
            {
                var worksheet = workbook.Worksheet(worksheetName); // Worksheet seçildi.

                json = GenerateJson(worksheet);
            }
            else
            {
                List<string> months = new() { "Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık" };
                IXLWorksheet[] worksheets = new IXLWorksheet[12];

                var worksheet = workbook.Worksheet("Ocak");

                // Veri okuma işlemi
                int firstRow = 1;
                int lastRow = 10;
                int firstColumn = 1;
                int lastColumn = 5;

                for (int row = firstRow; row <= lastRow; row++)
                {
                    for (int col = firstColumn; col <= lastColumn; col++)
                    {
                        // Hücrenin değerini alın
                        var cellValue = worksheet.Cell(row, col).Value.ToString();

                        // Hücre değerini ekrana yazdırın
                        Console.WriteLine("Row {0}, Column {1}: {2}", row, col, cellValue);
                    }
                }

            }

            File.WriteAllText($"{outputFileName}.json", json);
        }
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
using Excel_To_Json;

//ExcelToJson.ToJson("gcd_info.xlsx", "gcd_info", "result");
ExcelToJson.ToJson("LoadSheet.xlsx", "denemeLoadSheet"); // exceli uygun json formatta düzenleyip yeni bir json dosyası oluşturur.

//JsonFormatter.JsonFormat("denemeLoadSheet.json", "newFileName");
//JsonFormatter.ConvertToRowLabelsGCD("gcdinfo.json", "newgcdinfo");
//var result = JsonFormatter.GetJsonValues("newLoadSheet.json", "Ocak", "ABJIST");
//Console.WriteLine(result);
//JsonFormatter.ConvertJsonValues("ÜLKE ÇEVİRİ LİSTESİ.json");
Console.WriteLine("tamamdır.");

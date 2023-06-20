using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Excel_To_Json;

public static class JsonFormatter
{

    public static void ConvertToRowLabelsGCD(string filePath, string newFileName)
    {
        JArray updatedData = new JArray();
        JObject objectData = JObject.Parse(File.ReadAllText(filePath));
        _ = objectData.TryGetValue("Result", out var gcdData);
        //JArray data = JArray.Parse(gcdData.);


        foreach (var item in gcdData)
        {
            string org = item["ORG"].ToString();
            string dst = item["DST"].ToString();
            int gcd = Convert.ToInt32(item["GCD"]);

            string rowLabels = org + dst;

            JObject updatedItem = new JObject();
            updatedItem["Row Labels"] = rowLabels;
            updatedItem["GCD"] = gcd;

            updatedData.Add(updatedItem);
        }

        string updatedJson = JsonConvert.SerializeObject(updatedData, Formatting.Indented);
        File.WriteAllText($"{newFileName}.json", updatedJson);
    }

    public static void JsonFormat(string filePath, string newFileName)
    {
        JObject data = JObject.Parse(File.ReadAllText(filePath));
        JObject updatedData = new JObject();

        foreach (var month in data)
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

        File.WriteAllText($"{newFileName}.json", updatedData.ToString());
        Console.WriteLine("Updated JSON file has been created.");
    }

    public static string GetJsonValues(string filePath, string month, string rowLabels)
    {
        JObject data = JObject.Parse(File.ReadAllText(filePath));
        JArray subRows = GetSubRows(data, month, rowLabels);

        if (subRows != null)
        {
            string result = "[\n";
            Console.WriteLine("SubRows for {0}, {1}:", month, rowLabels);
            foreach (var subRow in subRows)
            {
                result += subRow.ToString() + ",\n";
            }
            result += "]";

            return result;
        }
        else
        {
            return string.Empty;
        }
    }

    public static void ConvertJsonValues(string filePath)
    {
        string json = File.ReadAllText(filePath);

        JArray countriesArray = JArray.Parse(json);

        Country[] countries = countriesArray.ToObject<Country[]>();

        ValuesObject[] trValues = new ValuesObject[countries.Length];
        ValuesObject[] enValues = new ValuesObject[countries.Length];

        for (int i = 0; i < countries.Length; i++)
        {
            trValues[i] = new ValuesObject
            {
                Key = countries[i].ENGLISH,
                Value = countries[i].TURKISH
            };

            enValues[i] = new ValuesObject
            {
                Key = countries[i].ENGLISH,
                Value = countries[i].ENGLISH
            };
        }

        CountryValues trCountryValues = new CountryValues
        {
            Values = trValues
        };

        CountryValues enCountryValues = new CountryValues
        {
            Values = enValues
        };

        string trJson = JsonConvert.SerializeObject(trCountryValues, Formatting.Indented);
        string enJson = JsonConvert.SerializeObject(enCountryValues, Formatting.Indented);

        File.WriteAllText("countries_tr.json", trJson);
        File.WriteAllText("countries_en.json", enJson);
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

    private static JArray GetSubRows(JObject data, string month, string rowLabels)
    {
        if (data.TryGetValue(month, out var monthData))
        {
            foreach (var item in monthData)
            {
                if (item["Row Labels"].ToString() == rowLabels && item["SubRows"] is JArray subRows)
                {
                    return subRows;
                }
            }
        }

        return null;
    }

    
}


public class Country
{
    public string ENGLISH { get; set; }
    public string TURKISH { get; set; }
}

public class ValuesObject
{
    public string Key { get; set; }
    public string Value { get; set; }
}

public class CountryValues
{
    public ValuesObject[] Values { get; set; }
}
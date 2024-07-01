using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System.Data.Common;
using System.Text;
using System.Data.OleDb;
using ExcelDataReader;
using Newtonsoft.Json.Schema.Generation;
using Newtonsoft.Json.Schema;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Text.Json.Nodes;


namespace FileProcessingAPI.Helpers;

public class ConvertExcelToJson
{
    public static async Task<string> ExcelToJson(string pathToExcel, string sheetName)
    {
        //This connection string works if you have Office 2007+ installed and your 
        //data is saved in a .xlsx file
        var connectionString = string.Format(@"
            Provider=Microsoft.ACE.OLEDB.12.0;
            Data Source={0};
            Extended Properties=""Excel 12.0 Xml;HDR=YES""
        ", pathToExcel);

        string json = string.Empty;

        //Creating and opening a data connection to the Excel sheet 
        using (var conn = new OleDbConnection(connectionString))
        {
            conn.Open();

            var cmd = conn.CreateCommand();
            cmd.CommandText = string.Format(
                @"SELECT * FROM [{0}$]",
                sheetName
                );

            using (var rdr = cmd.ExecuteReader())
            {
                //LINQ query - when executed will create anonymous objects for each row
                var query =
                                      (from DbDataRecord row in rdr
                                       select row).Select(x =>
                                       {
                                           //dynamic item = new ExpandoObject();
                                           Dictionary<string, object> item = new Dictionary<string, object>();
                                           for (int i = 0; i < x.FieldCount; i++)
                                               item.Add(rdr.GetName(i), x[i]);
                                           return item;
                                       });

                //Generates JSON from the LINQ query
                json = JsonConvert.SerializeObject(query); //, Formatting.Indented


            }
        }

        return json;

    }

    //https://stackoverflow.com/questions/57378535/read-exel-files-dynamically-not-depending-on-rows-and-write-json-c-sharp
    public static async Task<string> ExcelToJson(string inFilePath)
    {
        var json = string.Empty;

        using (var inFile = System.IO.File.Open(inFilePath, FileMode.Open, FileAccess.Read))

        using (var reader = ExcelReaderFactory.CreateReader(inFile, new ExcelReaderConfiguration()
        { FallbackEncoding = Encoding.GetEncoding(1252) }))
        {
            var ds = reader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true
                }
            });
            var table = ds.Tables[0];
            //json = JsonConvert.SerializeObject(table, Formatting.Indented);

            //var jsonSerializerSettings = new JsonSerializerSettings
            //{
            //    StringEscapeHandling = StringEscapeHandling.EscapeNonAscii
            //};
            //json = JsonConvert.SerializeObject(table, Formatting.Indented, jsonSerializerSettings);
            json = JsonConvert.SerializeObject(table, Formatting.Indented);
            json = RemoveSpecialCharactersFromJson(json);

        }

        return json;
    }

    //https://www.newtonsoft.com/jsonschema
    //https://www.newtonsoft.com/jsonschema/help/html/CustomJsonValidator.htm
    public static async Task<bool> ValidateJsonSchemaArray(string json, Type type)
    {
        if (string.IsNullOrEmpty(json)) // return false if null
            return false;
        JSchemaGenerator generator = new JSchemaGenerator();
        JSchema schema = generator.Generate(type);
        JArray jsonArray = JArray.Parse(json);
        if (jsonArray.Count == 0) // Return False if empty array
            return false;
        IList<string> messages;
        return jsonArray.IsValid(schema, out messages);
    }


    static string RemoveSpecialCharactersFromJson(string jsonString)
    {
        if (CheckJsonType(jsonString) == "JSON Array")
        {
            // Deserialize the JSON string into a JObject
            JArray jsonObject = JsonConvert.DeserializeObject<JArray>(jsonString);

            // Recursively process the JObject to remove special characters
            JArray cleanedJsonObject = (JArray)CleanJson(jsonObject);

            // Serialize the cleaned JObject back to a JSON string
            return JsonConvert.SerializeObject(cleanedJsonObject, Formatting.Indented);
        }else if (CheckJsonType(jsonString) == "JSON Object")
        {
            // Deserialize the JSON string into a JObject
            JObject jsonObject = JsonConvert.DeserializeObject<JObject>(jsonString);

            // Recursively process the JObject to remove special characters
            JObject cleanedJsonObject = (JObject)CleanJson(jsonObject);

            // Serialize the cleaned JObject back to a JSON string
            return JsonConvert.SerializeObject(cleanedJsonObject, Formatting.Indented);
        }else
            return string.Empty; // Error

    }

    static JToken CleanJson(JToken token)
    {
        if (token is JObject)
        {
            JObject obj = (JObject)token;
            JObject cleanedObj = new JObject();
            foreach (var property in obj.Properties())
            {
                cleanedObj.Add(property.Name, CleanJson(property.Value));
            }
            return cleanedObj;
        }
        else if (token is JArray)
        {
            JArray array = (JArray)token;
            JArray cleanedArray = new JArray();
            foreach (var item in array)
            {
                cleanedArray.Add(CleanJson(item));
            }
            return cleanedArray;
        }
        else if (token is JValue)
        {
            JValue value = (JValue)token;
            if (value.Type == JTokenType.String)
            {
                string cleanedValue = RemoveSpecialCharacters(value.ToString());
                return new JValue(cleanedValue);
            }
            return value;
        }
        return token;
    }

    static string RemoveSpecialCharacters(string input)
    {
        // Define a regular expression to match special characters
        // Adjust the pattern to match the specific characters you want to remove
        //Regex regex = new Regex("[^a-zA-Z0-9{}:,\"]+"); // Original Example
        Regex regex = new Regex("[^a-zA-Z0-9{}:,']+");

        // Replace special characters with an empty string
        return regex.Replace(input, "");
    }

    static string CheckJsonType(string jsonString)
    {
        try
        {
            // Parse the JSON string into a JToken
            JToken token = JToken.Parse(jsonString);

            // Check the type of the JToken
            if (token.Type == JTokenType.Object)
            {
                return "JSON Object";
            }
            else if (token.Type == JTokenType.Array)
            {
                return "JSON Array";
            }
            else
            {
                return "Unknown JSON type";
            }
        }
        catch (JsonReaderException)
        {
            return "Invalid JSON string";
        }
    }
}



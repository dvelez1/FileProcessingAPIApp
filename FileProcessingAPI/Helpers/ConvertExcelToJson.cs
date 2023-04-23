using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System.Data.Common;
using System.Text;
using System.Data.OleDb;
using ExcelDataReader;
using Newtonsoft.Json.Schema.Generation;
using Newtonsoft.Json.Schema;
using System.Text.Json;
using JsonSerializer = System.Text.Json.JsonSerializer;

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
            json = JsonConvert.SerializeObject(table, Formatting.Indented);
        }

        return json;
    }

    //https://www.newtonsoft.com/jsonschema
    //https://www.newtonsoft.com/jsonschema/help/html/CustomJsonValidator.htm
    public static async Task<bool> ValidateJsonSchemaArray(string json, Type type)
    {
        JSchemaGenerator generator = new JSchemaGenerator();
        JSchema schema = generator.Generate(type);
        JArray jsonArray = JArray.Parse(json);
        IList<string> messages;
        return jsonArray.IsValid(schema, out messages);
    }

}

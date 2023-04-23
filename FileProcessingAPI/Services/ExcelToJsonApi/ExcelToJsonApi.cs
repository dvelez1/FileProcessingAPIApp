using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Mvc;
using System.Runtime.CompilerServices;
using System.Xml.Linq;

namespace FileProcessingAPI.Service.Proccess1Api;

public static class ExcelToJsonApi
{

    public static void ConfigureExcelToJsonApi(this WebApplication app)
    {
       app.MapGet(pattern: "/FileProcessing/{excelFilePath}/ReadExcelAndConvertToJson", ReadExcelAndConvertToJson);
       app.MapGet(pattern: "/FileProcessing/{excelFilePath}/{sheetName}/ReadExcelAndConvertToJsonSecondAlternative", ReadExcelAndConvertToJssonPathAndSheetName);
       app.MapGet(pattern: "/FileProcessing/GetEmployees", GetEmployees);
       app.MapPost(pattern: "/FileProcessing/InsertExcelDataSetIntoEmployee/{excelFilePath}", InsertExcelIntoEmployee);
    }

    private static async Task<IResult> ReadExcelAndConvertToJson(string excelFilePath)
    {
        try
        {
            return Results.Ok(await Helpers.ConvertExcelToJson.ExcelToJson(excelFilePath));
        }
        catch (Exception ex)
        {
            return Results.Problem(ex.Message);
        }

    }
    private static async Task<IResult> ReadExcelAndConvertToJssonPathAndSheetName(string excelFilePath,string sheetName)
    {
        try
        {
            return Results.Ok(await Helpers.ConvertExcelToJson.ExcelToJson(excelFilePath,sheetName));
        }
        catch (Exception ex)
        {
            return Results.Problem(ex.Message);
        }

    }
    private static async Task<IResult> InsertExcelIntoEmployee(string excelFilePath, IEmployeeData data)
    {
        try
        {
            var json = await Helpers.ConvertExcelToJson.ExcelToJson(excelFilePath);
            if (await Helpers.ConvertExcelToJson.ValidateJsonSchemaArray(json, typeof(List<EmployeeModel>)))
            {
                await data.InsertEmployeeList(json);
                return Results.Ok("Success Transaction!");
            } else
                return Results.Problem($"Json Schema Validation Failed!");
        }
        catch (Exception ex)
        {
            return Results.Problem(ex.Message);
        }

    }
    private static async Task<IResult> GetEmployees(IEmployeeData data)
    {
        try
        {
            return Results.Ok(await data.GetEmployees());
        }
        catch (Exception ex)
        {
            return Results.Problem(ex.Message);
        }
    }

}

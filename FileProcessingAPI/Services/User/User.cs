using FileProcessingAPI.Helpers;
using Microsoft.AspNetCore.Routing.Constraints;

namespace FileProcessingAPI.Service.User;

public static class User
{
    // Only will be exposed ConfigureApi, because the methods are private!
    public static void ConfigureApi(this WebApplication app)
    {
        // API endpoint mapping
        app.MapGet(pattern: "/Users", GetUsers);
        app.MapGet(pattern: "/Users/{id}", GetUser);
        app.MapPost(pattern: "/Users", InsertUser);
        app.MapPut(pattern: "/Users", UpdateUser);
        app.MapDelete(pattern: "/Users", DeleteUser);
        app.MapGet(pattern: "/Users/UsersAllToExcel", UsersAllToExcel);
    }

    private static async Task<IResult> GetUsers(IUserData data)
    {
        try
        {
            return Results.Ok(await data.GetUsers());
        }
        catch (Exception ex)
        {
            return Results.Problem(ex.Message);
        }
    }

    private static async Task<IResult> GetUser(int id, IUserData data)
    {
        try
        {
            var results = await data.GetUser(id);
            if (results == null) return Results.NotFound();
            return Results.Ok(results);
        }
        catch (Exception ex)
        {
            return Results.Problem(ex.Message);
        }
    }

    private static async Task<IResult> InsertUser(UserModel user, IUserData data)
    {
        try
        {
            await data.InsertUser(user);
            return Results.Ok();
        }
        catch (Exception ex)
        {
            return Results.Problem(ex.Message);
        }
    }

    private static async Task<IResult> UpdateUser(UserModel user, IUserData data)
    {
        try
        {
            await data.UpdateUser(user);
            return Results.Ok();
        }
        catch (Exception ex)
        {
            return Results.Problem(ex.Message);
        }
    }

    private static async Task<IResult> DeleteUser(int id, IUserData data)
    {
        try
        {
            await data.DeleteUser(id);
            return Results.Ok(data);
        }
        catch (Exception ex)
        {
            return Results.Problem(ex.Message);
        }
    }

    private static async Task<IResult> UsersAllToExcel(string filePath, string fileName, IUserData data)
    {
        try
        {
            await UserToExcelMethod(filePath, fileName, data);
            return Results.Ok("Excel download Success!");
        }
        catch (Exception ex)
        {
            return Results.Problem(ex.Message);
        }
    }


    // Only for Demo purpose. The location of this method is not correct
    private static async Task UserToExcelMethod(string filePath, string fileName, IUserData data)
    {
        var modelList = await data.GetUsers();
        //Parallel.Invoke(
        //    () => ConvertListToExcel.GenerateExcel(ConvertListToExcel.ConvertToDataTable(modelList as List<UserModel>), filePath, fileName),
        //    () => ConvertListToExcel.GenerateExcel(ConvertListToExcel.ConvertToDataTable(modelList as List<UserModel>), filePath, fileName + "2"),
        //    () => ConvertListToExcel.GenerateExcel(ConvertListToExcel.ConvertToDataTable(modelList as List<UserModel>), filePath, fileName + "3")
        //    );
        
        ConvertListToExcel.GenerateExcel(ConvertListToExcel.ConvertToDataTable(modelList as List<UserModel>), filePath, fileName);
        ConvertListToExcel.SaveToCsv(ConvertListToExcel.ConvertToDataTable(modelList as List<UserModel>), filePath, fileName);
    }


}

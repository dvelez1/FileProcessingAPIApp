using DataAccess.DbAccess;
using DataAccess.Models;
using DataAccess.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace DataAccess.Data;

public class UserData : IUserData
{
    private readonly ISqlDataAccess _db;

    public UserData(ISqlDataAccess db) => _db = db;


    public Task<IEnumerable<UserModel>> GetUsers() =>
         _db.LoadData<UserModel, dynamic>(storeProcedure: "dbo.spUser_GetAll", new { });


    public async Task<UserModel?> GetUser(int id)
    {
        var results = await _db.LoadData<UserModel, dynamic>(
            storeProcedure: "dbo.spUser_Get",
            new { id = id });

        return results.FirstOrDefault();
    }

    public Task InsertUser(UserModel user) =>
        _db.SaveData(storeProcedure: "dbo.spUser_Insert", new { user.FirstName, user.LastName });

    public Task UpdateUser(UserModel user) =>
        _db.SaveData(storeProcedure: "dbo.spUser_Update", user);

    public Task DeleteUser(int id) =>
        _db.SaveData(storeProcedure: "dbo.spUser_Delete", new { Id = id });


    public Task<int> InsertUserWithDynamicParameters(UserModel user) => _db.SaveDataByDynamicParameter("dbo.spUser_Insert",
             DynamicDapperParametersMapper.DynamicParametersMapper(new { FirstName= user.FirstName, LastName= user.LastName }));

    /// <summary>
    /// Example with retun value
    /// </summary>
    /// <param name="id"></param>
    /// <returns></returns>
    public async Task<Tuple<UserModel?, bool>> GetUserWithReturValue(int id)
    {
        var (list,result) = await _db.LoadDataWithRetunValue<UserModel>(
                storeProcedure: "dbo.spUser_Get",
                    DynamicDapperParametersMapper.DynamicParametersMapper(new { id = id }));
        
        return new Tuple<UserModel?, bool>(list.FirstOrDefault(), result>0);
    }
}

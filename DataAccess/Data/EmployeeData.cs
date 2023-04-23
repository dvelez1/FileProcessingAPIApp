using DataAccess.DbAccess;
using DataAccess.Models;


namespace DataAccess.Data;

public class EmployeeData : IEmployeeData
{
    private readonly ISqlDataAccess _db;

    public EmployeeData(ISqlDataAccess db) => _db = db;

    public Task InsertEmployeeList(string json) =>
        _db.SaveData(storeProcedure: "dbo.spEmployeeList_Insert", new { JSONCustomer = json });

    public Task<IEnumerable<EmployeeModel>> GetEmployees() =>
        _db.LoadData<EmployeeModel, dynamic>(storeProcedure: "dbo.spEmployee_GetAll", new { });

}

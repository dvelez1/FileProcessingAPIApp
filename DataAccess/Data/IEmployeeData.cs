using DataAccess.Models;

namespace DataAccess.Data
{
    public interface IEmployeeData
    {
        Task<IEnumerable<EmployeeModel>> GetEmployees();
        Task InsertEmployeeList(string json);
    }
}
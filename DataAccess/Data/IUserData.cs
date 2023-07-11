using DataAccess.Models;

namespace DataAccess.Data
{
    public interface IUserData
    {
        Task DeleteUser(int id);
        Task<UserModel?> GetUser(int id);
        Task<IEnumerable<UserModel>> GetUsers();
        Task InsertUser(UserModel user);
        Task UpdateUser(UserModel user);

        Task<int> InsertUserWithDynamicParameters(UserModel user);
        Task<Tuple<UserModel?, bool>> GetUserWithReturValue(int id);
    }
}
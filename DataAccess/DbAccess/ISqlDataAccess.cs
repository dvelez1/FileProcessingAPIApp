using Dapper;

namespace DataAccess.DbAccess
{
    public interface ISqlDataAccess
    {
        Task<IEnumerable<T>> LoadData<T, U>(string storeProcedure, U parameters, string connectionId = "Default");
        Task SaveData<T>(string storeProcedure, T parameters, string connectionId = "Default");

        #region Methods with return value (CRUD)
        Task<int> SaveDataByDynamicParameter(string storeProcedure, DynamicParameters parameters, string connectionId = "Default");
        Task<Tuple<IEnumerable<T>, int>> LoadDataWithRetunValue<T>(string storeProcedure, DynamicParameters parameters, string connectionId = "Default");
        #endregion

    }
}
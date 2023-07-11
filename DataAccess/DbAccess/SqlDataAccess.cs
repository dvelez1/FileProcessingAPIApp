using Microsoft.Extensions.Configuration;
using Dapper;
using System.Data;
using System.Data.SqlClient;



namespace DataAccess.DbAccess;

public class SqlDataAccess : ISqlDataAccess
{
    private readonly IConfiguration _config;

    public SqlDataAccess(IConfiguration config)
    {
        _config = config;
    }

    public async Task<IEnumerable<T>> LoadData<T, U>(
        string storeProcedure,
        U parameters,
        string connectionId = "Default")
    {
        using IDbConnection connection = new SqlConnection(_config.GetConnectionString(connectionId));

        return await connection.QueryAsync<T>(storeProcedure, parameters,
            commandType: CommandType.StoredProcedure);

    }

    public async Task SaveData<T>(
        string storeProcedure,
        T parameters,
        string connectionId = "Default")
    {
        using IDbConnection connection = new SqlConnection(_config.GetConnectionString(connectionId));

        await connection.ExecuteAsync(storeProcedure, parameters,
            commandType: CommandType.StoredProcedure);

    }

    #region Methods with return value (CRUD)

    public async Task<int> SaveDataByDynamicParameter(string storeProcedure, DynamicParameters parameters, string connectionId = "Default")
    {
        using IDbConnection connection = new SqlConnection(_config.GetConnectionString(connectionId));

        await connection.ExecuteAsync(storeProcedure, parameters, commandType: CommandType.StoredProcedure);
        return parameters.Get<int>("@ReturnVal");
    }

    public async Task<Tuple<IEnumerable<T>, int>> LoadDataWithRetunValue<T>(
        string storeProcedure,
        DynamicParameters parameters,
        string connectionId = "Default")
    {
        using IDbConnection connection = new SqlConnection(_config.GetConnectionString(connectionId));

        var list = await connection.QueryAsync<T>(storeProcedure, parameters,
            commandType: CommandType.StoredProcedure);

        int transactionResult = parameters.Get<int>("@ReturnVal");

        return new Tuple<IEnumerable<T>, int>(list, transactionResult);

    }

    #endregion

}

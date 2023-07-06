﻿using Dapper;

namespace DataAccess.DbAccess
{
    public interface ISqlDataAccess
    {
        Task<IEnumerable<T>> LoadData<T, U>(string storeProcedure, U parameters, string connectionId = "Default");
        Task SaveData<T>(string storeProcedure, T parameters, string connectionId = "Default");

        Task<int> SaveDataByDynamicParameter<T>(string storeProcedure, DynamicParameters parameters, string connectionId = "Default");
    }
}
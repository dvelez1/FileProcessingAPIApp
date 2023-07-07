using Dapper;
using Newtonsoft.Json.Serialization;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataAccess.Utilities;
public static class DynamicDapperParametersMapper
{
    private static readonly Dictionary<Type, DbType> TypeMap = new()
    {
        {typeof(byte),DbType.Byte},
        {typeof(sbyte),DbType.SByte},
        {typeof(short),DbType.Int16},
        {typeof(ushort),DbType.UInt16},
        {typeof(int),DbType.Int32},
        {typeof(uint),DbType.UInt32},
        {typeof(long),DbType.Int64},
        {typeof(ulong),DbType.UInt64},
        {typeof(float),DbType.Single},
        {typeof(double),DbType.Double},
        {typeof(decimal),DbType.Decimal},
        {typeof(bool),DbType.Boolean},
        {typeof(string),DbType.String},
        {typeof(char),DbType.StringFixedLength},
        {typeof(Guid),DbType.Guid},
        {typeof(DateTime),DbType.DateTime},
        {typeof(DateTimeOffset),DbType.DateTimeOffset},
        {typeof(byte[]),DbType.Binary},
        {typeof(byte?),DbType.Byte},
        {typeof(sbyte?),DbType.SByte},
        {typeof(short?),DbType.Int16},
        {typeof(ushort?),DbType.UInt16},
        {typeof(int?),DbType.Int32},
        {typeof(uint?),DbType.UInt32},
        {typeof(long?),DbType.Int64},
        {typeof(ulong?),DbType.UInt64},
        {typeof(float?),DbType.Single},
        {typeof(double?),DbType.Double},
        {typeof(decimal?),DbType.Decimal},
        {typeof(bool?),DbType.Boolean},
        {typeof(char?),DbType.StringFixedLength},
        {typeof(Guid?),DbType.Guid},
        {typeof(DateTime?),DbType.DateTime},
        {typeof(DateTimeOffset?),DbType.DateTimeOffset}
    };

    public static DynamicParameters DynamicParametersMapper(this object obj)
    {
        var parameters = new DynamicParameters();

        var resolve = new DefaultContractResolver() { NamingStrategy = new SnakeCaseNamingStrategy() };
        foreach (var prop in obj.GetType().GetProperties())
        {
            //var pName = resolve.GetResolvedPropertyName(prop.Name); // Convert to FirstName to first_name
            var pName = prop.Name;
            parameters.Add($"@{pName}", prop.GetValue(obj, null), TypeMap[prop.PropertyType], ParameterDirection.Input);
        }
        // parameters.Add("ReturnVal", dbType: DbType.Int32, direction: ParameterDirection.Output); Used with OUTPUT Parameter on SP. Example: ,@ReturnVal int OUTPUT
        parameters.Add("ReturnVal", dbType: DbType.Int32, direction: ParameterDirection.ReturnValue);

        return parameters;
    }


}

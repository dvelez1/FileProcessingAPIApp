using DataAccess.DbAccess;
using FileProcessingAPI.Service.Proccess1Api;
using FileProcessingAPI.Service.User;

var builder = WebApplication.CreateBuilder(args);

System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance); // Used Convert Excel to Json Method
Dapper.DefaultTypeMap.MatchNamesWithUnderscores = true; // Remove Underscore (_) for automatic Object Mapping

// Add services to the container.
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
#region Services
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
#endregion

#region Interfaces Registration
//Add My Interfaces from DataAccess Library (Register the services)
builder.Services.AddSingleton<ISqlDataAccess, SqlDataAccess>();
builder.Services.AddSingleton<IUserData, UserData>();
builder.Services.AddScoped<IEmployeeData, EmployeeData>();
#endregion

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();

#region endpoint mapping
app.ConfigureApi(); 
app.ConfigureExcelToJsonApi();
#endregion

app.Run();

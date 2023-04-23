CREATE TABLE [dbo].[Employee]
(
	[employee_id] INT NOT NULL PRIMARY KEY IDENTITY, 
    [full_name] NVARCHAR(250) NULL, 
    [job_title] NVARCHAR(250) NULL, 
    [department] NVARCHAR(250) NULL, 
    [business_unit] NVARCHAR(250) NULL, 
    [gender] NVARCHAR(50) NULL
)

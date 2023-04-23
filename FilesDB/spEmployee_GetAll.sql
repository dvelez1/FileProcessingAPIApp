CREATE PROCEDURE [dbo].[spEmployee_GetAll]
AS
BEGIN
 Select
   employee_id,
    full_name,
    job_title,
    department, 
    business_unit,gender from dbo.Employee
END

CREATE PROCEDURE [dbo].[spEmployeeList_Insert]
	@JSONCustomer nvarchar(max)
	AS
BEGIN

DELETE FROM dbo.Employee;

INSERT INTO dbo.Employee (
  employee_id,
    full_name,
    job_title,
    department, 
    business_unit,gender)

SELECT employee_id,
    full_name,
    job_title,
    department, 
    business_unit,gender
FROM  
 OPENJSON ( @JSONCustomer )  
WITH (   
              employee_id   nvarchar(200) '$.employee_id' ,  
              full_name     nvarchar(200) '$.full_name',  
              job_title nvarchar(200)     '$.job_title',  
              department nvarchar(200)   '$.department',
              business_unit nvarchar(200) '$.business_unit',
              gender nvarchar(200)       '$.gender'
 ) 


END

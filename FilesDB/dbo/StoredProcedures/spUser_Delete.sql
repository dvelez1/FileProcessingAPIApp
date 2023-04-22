CREATE PROCEDURE [dbo].[spUser_Delete]
	@Id int
AS
BEGIN
	DELETE
	FROM dbo.[User]
	where Id = @Id;
END

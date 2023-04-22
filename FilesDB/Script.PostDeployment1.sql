IF NOT EXISTS (SELECT 1 FROM dbo.[User])
BEGIN
    INSERT INTO dbo.[User] (FirstName, LastName)
    values ('Dennis', 'Velez'), ('Ramon','Perez'), ('Nery','Nelson');
END
/* Example implementation on MS SQL Server using xp_cmdshell */

IF OBJECT_ID('dbo.Out_to_Excel','P') IS NOT NULL DROP PROC dbo.Out_to_Excel
GO

CREATE PROC [dbo].[Out_to_Excel] (
@save_to_folder NVARCHAR(300) = N'.', 
@fileName NVARCHAR(125),
@server NVARCHAR(125) = N'localhost',
@database NVARCHAR(125) = N'master', 
@sqlfile NVARCHAR(125),
@worksheetname NVARCHAR(30) = N'Sheet1',
@append_date BIT = 0
)
AS

DECLARE @cmd VARCHAR(8000)

-- Additional wrappers getting added to directory pathing due to the way xp_cmdshell parses file paths
SET @sqlfile = '''''' + @sqlfile + ''''''
SET @save_to_folder = '''''' + @save_to_folder + ''''''
SET @fileName = '''''' + @fileName + ''''''

BEGIN TRY

	SET @cmd = N'C:\SqlToExcel.ps1 -save_to_folder "' + @save_to_folder + N'" -fileName "' + @fileName
				+ N'" -server "' + @server + N'" -database ' + @database +  N' -sqlfile "' + @sqlfile
				+ N'" -worksheetname "' + @worksheetname + N'" -append_date ' + (CASE WHEN @append_date = 0 THEN N'0' ELSE N'1' END)

	SET @cmd = '''powershell ' + @cmd + '''' 
	SELECT @cmd
	EXEC ('xp_cmdshell ' + @cmd)
	--
END TRY

BEGIN CATCH
			SELECT ERROR_NUMBER() AS ErrorNumber, ERROR_SEVERITY() AS ErrorSeverity, ERROR_STATE() AS ErrorState,
			ERROR_PROCEDURE() AS ErrorProcedure, ERROR_LINE() AS ErrorLine,	ERROR_MESSAGE() AS ErrorMessage;
END CATCH

GO



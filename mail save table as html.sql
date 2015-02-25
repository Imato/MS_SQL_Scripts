/*
Formate same t-sql data table as html table

*/

ALTER proc [dbo].[prc_save_table_as_html]
	@HTML varchar(max) out, 
    @PCWrite varchar(1000) = NULL,
    @DBFetch varchar(max),
    @DBWhere varchar(max) = NULL,
    @DBThere varchar(max) = NULL,
    @DBUltra bit = 1,
    @TableStyle varchar(1000) = 'border-width: thin; border-spacing: 2px; border-style: solid; border-color: gray; border-collapse: collapse;',
    @CSS varchar(max) = NULL,
    @Header bit = 1 -- Output header. Default is 1.
	
AS

SET NOCOUNT ON

DECLARE @Return int
DECLARE @Retain int
DECLARE @Status int

SET @Status = 0
set @HTML=''

DECLARE @TPre varchar(10)

DECLARE @TDo3 tinyint
DECLARE @TDo4 tinyint

SET @TPre = ''

SET @TDo3 = LEN(@TPre)
SET @TDo4 = LEN(@TPre) + 1

DECLARE @DBAE varchar(250)
DECLARE @Task varchar(max)
DECLARE @Bank varchar(max)
DECLARE @Cash varchar(max)
DECLARE @Risk varchar(max)
DECLARE @Next varchar(max)
DECLARE @Save varchar(max)
DECLARE @Work varchar(max)
DECLARE @Wish varchar(max)
DECLARE @hString varchar(max)
DECLARE @tString varchar(max)
DECLARE @fString varchar(50)

DECLARE @Name varchar(100)
DECLARE @Same varchar(100)

DECLARE @Rank smallint
DECLARE @Kind varchar(20)
DECLARE @Mask bit
DECLARE @Bond bit
DECLARE @Size int
DECLARE @Wide smallint
DECLARE @More smallint

DECLARE @DBAI varchar(max)
DECLARE @DBAO varchar(max)
DECLARE @DBAU varchar(max)

DECLARE @Fuse int
DECLARE @File int

DECLARE @FuseO int
DECLARE @FileO int
DECLARE @TmpPathObj int
DECLARE @TmpPath varchar(127)
DECLARE @TmpFile varchar(127)
DECLARE @TmpFilename varchar(1000)

DECLARE @HeaderString varchar(max)
DECLARE @sHeaderString varchar(max)
DECLARE @HeaderDone int

SET @DBAE = '##SaveFile' + RIGHT(CONVERT(varchar(10),@@SPID+100000),5)

SET @Task = 'IF EXISTS (SELECT * FROM tempdb.dbo.sysobjects WHERE name = ' + CHAR(39) + @DBAE + CHAR(39) + ') DROP TABLE ' + @DBAE
EXECUTE (@Task)

SET @Bank = @TPre + @DBFetch

IF NOT EXISTS (SELECT * FROM sysobjects WHERE RTRIM(type) = 'U' AND name = @Bank)
BEGIN
	SET @Bank = CASE WHEN upper(LEFT(LTRIM(@DBFetch),6)) = 'SELECT' THEN '(' + @DBFetch + ')' ELSE @DBFetch END
	SET @Bank = REPLACE(@Bank,         CHAR(94),CHAR(39))
	SET @Bank = REPLACE(@Bank,CHAR(45)+CHAR(45),CHAR(32))
	SET @Bank = REPLACE(@Bank,CHAR(47)+CHAR(42),CHAR(32))
END

IF @DBWhere IS NOT NULL
BEGIN
	SET @Cash = REPLACE(@DBWhere,'WHERE'       ,CHAR(32))
	SET @Cash = REPLACE(@Cash,         CHAR(94),CHAR(39))
	SET @Cash = REPLACE(@Cash,CHAR(45)+CHAR(45),CHAR(32))
	SET @Cash = REPLACE(@Cash,CHAR(47)+CHAR(42),CHAR(32))
END

IF @DBThere IS NOT NULL
BEGIN
	SET @Risk = REPLACE(@DBThere,'ORDER BY'    ,CHAR(32))
	SET @Risk = REPLACE(@Risk,         CHAR(94),CHAR(39))
	SET @Risk = REPLACE(@Risk,CHAR(45)+CHAR(45),CHAR(32))
	SET @Risk = REPLACE(@Risk,CHAR(47)+CHAR(42),CHAR(32))
END

SET @DBAI = ''
SET @DBAO = ''
SET @DBAU = ''

SET @Task = 'SELECT * INTO ' + @DBAE + ' FROM ' + @Bank + ' AS T WHERE 0 = 1'
IF @Status = 0 EXECUTE (@Task) SET @Return = @@ERROR
IF @Status = 0 SET @Status = @Return

DECLARE Fields CURSOR FAST_FORWARD FOR
SELECT C.name, C.colid, T.name, C.isnullable, C.iscomputed, C.length, C.prec, C.scale
FROM tempdb.dbo.sysobjects AS O
JOIN tempdb.dbo.syscolumns AS C
  ON O.id = C.id
JOIN tempdb.dbo.systypes AS T
  ON C.xusertype = T.xusertype
WHERE O.name = @DBAE
ORDER BY C.colid

SET @Retain = @@ERROR IF @Status = 0 SET @Status = @Retain

OPEN Fields

SET @Retain = @@ERROR IF @Status = 0 SET @Status = @Retain

FETCH NEXT FROM Fields INTO @Same, @Rank, @Kind, @Mask, @Bond, @Size, @Wide, @More
SET @Retain = @@ERROR IF @Status = 0 SET @Status = @Retain

-- Convert to character for header.
SET @HeaderString = ''
SET @sHeaderString = ''
declare @sql nvarchar(max)
declare @aHeader nvarchar(9)
declare @zHeader nvarchar(9)
set @aHeader = '<th>'
set @zHeader = '</th>'

WHILE @@FETCH_STATUS = 0 AND @Status = 0
BEGIN
	-- Build header.
	IF LEN(@HeaderString) > 0 SET @HeaderString = @HeaderString + '+lower(''<tr>'')' + '+ISNULL(''' + @Same + ''',SPACE(0))+' + 'lower(''</tr>'')+'
	IF LEN(@HeaderString) = 0 SET @HeaderString = '+lower(''<tr>'')' + '+ISNULL(''' + @Same + ''',SPACE(0))+' + 'lower(''</tr>'')+'
	IF LEN(@sHeaderString) > 0 SET @sHeaderString = @sHeaderString + @aHeader + ISNULL(@Same, SPACE(0)) + @zHeader
	IF LEN(@sHeaderString) = 0 SET @sHeaderString = @aHeader + ISNULL(@Same, SPACE(0)) + @zHeader

	IF @Kind IN ('char','varchar','nchar','nvarchar')
	BEGIN
		IF @Rank = 1 SET @DBAU = 'lower(''<td>'')' + '+ISNULL(CONVERT(varchar(250),' + @Same + '),SPACE(0))+' + 'lower(''</td>'')'
		IF @Rank > 1 SET @DBAU = @DBAU + '+lower(''<td>'')' + '+ISNULL(CONVERT(varchar(250),' + @Same + '),SPACE(0))+' + 'lower(''</td>'')'
	END

	IF @Kind IN ('bit','tinyint','smallint','int','bigint')
	BEGIN
		IF @Rank = 1 SET @DBAU = 'lower(''<td>'')' + '+ISNULL(CONVERT(varchar(250),' + @Same + '),SPACE(0))+' + 'lower(''</td>'')'
		IF @Rank > 1 SET @DBAU = @DBAU + '+lower(''<td>'')' + '+ISNULL(CONVERT(varchar(250),' + @Same + '),SPACE(0))+' + 'lower(''</td>'')'
	END

	IF @Kind IN ('numeric','decimal','money','smallmoney','float','real')
	BEGIN
		IF @Rank = 1 SET @DBAU = 'lower(''<td>'')' + '+ISNULL(CONVERT(varchar(250),' + @Same + '),SPACE(0))+' + 'lower(''</td>'')'
		IF @Rank > 1 SET @DBAU = @DBAU + '+lower(''<td>'')' + '+ISNULL(CONVERT(varchar(250),' + @Same + '),SPACE(0))+' + 'lower(''</td>'')'
	END

	IF @Kind IN ('uniqueidentifier')
	BEGIN
		IF @Rank = 1 SET @DBAU = 'lower(''<td>'')' + '+ISNULL(CONVERT(varchar(250),' + @Same + '),SPACE(0))+' + 'lower(''</td>'')'
		IF @Rank > 1 SET @DBAU = @DBAU + '+lower(''<td>'')' + '+ISNULL(CONVERT(varchar(250),' + @Same + '),SPACE(0))+' + 'lower(''</td>'')'
	END

	IF @Kind IN ('datetime','smalldatetime')
	BEGIN
		IF @Rank = 1 SET @DBAU = 'lower(''<td>'')' + '+ISNULL(CONVERT(varchar(250),' + @Same + '),SPACE(0))+' + 'lower(''</td>'')'
		IF @Rank > 1 SET @DBAU = @DBAU + '+lower(''<td>'')' + '+ISNULL(CONVERT(varchar(250),' + @Same + '),SPACE(0))+' + 'lower(''</td>'')'
	END

	FETCH NEXT FROM Fields INTO @Same, @Rank, @Kind, @Mask, @Bond, @Size, @Wide, @More
	SET @Retain = @@ERROR IF @Status = 0 SET @Status = @Retain
END

CLOSE Fields DEALLOCATE Fields

IF LEN(@DBAU) = 0 SET @DBAU = '*'

SET @DBAI = ' SELECT '
SET @DBAO = '   FROM ' + @Bank + ' AS T'
	+ CASE WHEN @DBWhere IS NULL THEN '' ELSE ' WHERE (' + @Cash + ') AND 0 = 0' END
	+ CASE WHEN @DBThere IS NULL THEN '' ELSE ' ORDER BY ' + @Risk END

IF LEN(ISNULL(@PCWrite,'*')) > 7 AND @DBUltra = 0
BEGIN
	IF @CSS IS NULL
	BEGIN
		SET @tString = ' select lower(''<html><body><table border='') + CHAR(34) + ''1'' + CHAR(34) + '' style='' + CHAR(34) + lower(''' + @TableStyle + ''') + CHAR(34) + ''>'' UNION ALL '
	END
	ELSE
	BEGIN
		SET @tString = ' select lower(''<html><head><style type="text/css">'') + lower(''' + @CSS + ''') + ''</style></head><body><table>'' UNION ALL '
	END
	SET @fString = ' UNION ALL select ''</table></body></html>'''
	SET @hString = ''
	IF @Header = 1
	BEGIN
		SET @hString = ' select ''<tr>' + @sHeaderString + '</tr>'' UNION ALL '
	END
	SET @Wish = 'set nocount on; USE ' + DB_NAME() + @tString + @hString + @DBAI + '''<tr>''+' + @DBAU + '+''</tr>''' + @DBAO + @fString
	-- SET @Work = 'bcp "' + @Wish + '" queryout "' + @PCWrite + '" -c -T' -- Query length of BCP is limited to only 1023 chars.
	-- Create SQL script file.
	IF @Status = 0 EXECUTE @Status = sp_OACreate 'Scripting.FileSystemObject', @FuseO OUTPUT
	IF @Status = 0 EXECUTE @Status = sp_OAGetProperty @FuseO, 'GetSpecialFolder(2)', @TmpPathObj OUTPUT
	IF @Status = 0 EXECUTE @Status = sp_OAGetProperty @TmpPathObj, 'Path', @TmpPath OUTPUT
	IF @Status = 0 EXECUTE @Status = sp_OAGetProperty @FuseO, 'GetTempName', @TmpFile OUTPUT
	SET @TmpFilename = @TmpPath + '\' + @TmpFile
	IF @Status = 0 EXECUTE @Status = sp_OAMethod @FuseO, 'CreateTextFile', @FileO OUTPUT, @TmpFilename, -1
	IF @Status <> 0 GOTO ABORT
	IF @Status = 0 EXECUTE @Status = sp_OAMethod @FileO, 'Write', NULL, @Wish
	IF @Status = 0 EXECUTE @Status = sp_OAMethod @FileO, 'Close'

	SET @Work = 'osql -i "' + @TmpFilename + '" -o "' + @PCWrite + '" -n -h-1 -w8000 -E'
	EXECUTE @Return = master.dbo.xp_cmdshell @Work, NO_OUTPUT
	SET @Retain = @@ERROR
	IF @Status = 0 SET @Status = @Retain
	IF @Status = 0 SET @Status = @Return

	EXECUTE @Status = sp_OAMethod @FuseO, 'DeleteFile', NULL, @TmpFilename
	EXECUTE @Status = sp_OADestroy @FuseO

	GOTO ABORT
END

IF LEN(ISNULL(@PCWrite,'*')) > 7
BEGIN
	IF @Status = 0 EXECUTE @Return = sp_OACreate 'Scripting.FileSystemObject', @Fuse OUTPUT
	SET @Retain = @@ERROR
	IF @Status = 0 SET @Status = @Retain
	IF @Status = 0 SET @Status = @Return

	IF @Status = 0 EXECUTE @Return = sp_OAMethod @Fuse, 'CreateTextFile', @File OUTPUT, @PCWrite, -1
	SET @Retain = @@ERROR
	IF @Status = 0 SET @Status = @Retain
	IF @Status = 0 SET @Status = @Return

	IF @Status <> 0 GOTO ABORT
END

SET @DBAI = 'DECLARE Records CURSOR GLOBAL FAST_FORWARD FOR' + @DBAI

IF @Status = 0 EXECUTE (@DBAI+@DBAU+@DBAO) SET @Return = @@ERROR
IF @Status = 0 SET @Status = @Return

OPEN Records
SET @Retain = @@ERROR IF @Status = 0 SET @Status = @Retain

FETCH NEXT FROM Records INTO @Next
SET @Retain = @@ERROR IF @Status = 0 SET @Status = @Retain

SET @HeaderDone = 0
WHILE @@FETCH_STATUS = 0 AND @Status = 0
BEGIN
	SET @Save = ''

	IF ISNULL(@File,0) = 0
	BEGIN
		-- Print header (TEXT).
		IF @HeaderDone = 0
		BEGIN
			IF @CSS IS NULL
			BEGIN
				set @HTML =  @HTML + '<table border="1" style="' + @TableStyle + '">' + CHAR(13) + CHAR(10)
			END
			ELSE
			BEGIN
				set @HTML =  @HTML + '<table>' + CHAR(13) + CHAR(10)
			END
			SET @HeaderDone = 1
		END
		IF @Header = 1
		BEGIN
			set @HTML =  @HTML + '<tr>' + @sHeaderString + '</tr>' + CHAR(13) + CHAR(10)
			SET @Header = 0
		END
		set @HTML =  @HTML + '<tr>' + @Next + '</tr>'
	END
	ELSE
	BEGIN
		-- Print header (FILE).
		IF @HeaderDone = 0
		BEGIN
			IF @CSS IS NULL
			BEGIN
				SET @Save = @Save + '<html><body><table border="1" style="' + @TableStyle + '">' + CHAR(13) + CHAR(10)
			END
			ELSE
			BEGIN
				SET @Save = @Save + '<html><head><style type="text/css">' + @CSS + '</style></head><body><table>' + CHAR(13) + CHAR(10)
			END
			SET @HeaderDone = 1
		END
		IF @Header = 1
		BEGIN
			SET @Save = @Save + '<tr>' + @sHeaderString + '</tr>' + CHAR(13) + CHAR(10)
			SET @Header = 0
		END

		-- Print the data.
		SET @Save = @Save + '<tr>' + @Next + '</tr>' + CHAR(13) + CHAR(10)
		IF @Status = 0 EXECUTE @Return = sp_OAMethod @File, 'Write', NULL, @Save
		IF @Status = 0 SET @Status = @Return
	END

	FETCH NEXT FROM Records INTO @Next
	SET @Retain = @@ERROR IF @Status = 0 SET @Status = @Retain
END

CLOSE Records DEALLOCATE Records

-- Print footer (TEXT).
IF ISNULL(@File,0) = 0
BEGIN
	set @HTML =  @HTML + '</table>' + CHAR(13) + CHAR(10)
END
ELSE
BEGIN
	SET @Save = '</table></body></html>' + CHAR(13) + CHAR(10)
	IF @Status = 0 EXECUTE @Return = sp_OAMethod @File, 'Write', NULL, @Save
END

-- Close.
IF ISNULL(@File,0) <> 0
BEGIN
	EXECUTE @Return = sp_OAMethod @File, 'Close', NULL
	IF @Status = 0 SET @Status = @Return

	EXECUTE @Return = sp_OADestroy @File
	IF @Status = 0 SET @Status = @Return

	EXECUTE @Return = sp_OADestroy @Fuse
	IF @Status = 0 SET @Status = @Return
END

ABORT: -- This label is referenced when OLE automation fails.

IF @Status = 1 OR @Status NOT BETWEEN 0 AND 50000 RAISERROR ('SaveTableAsHTML Windows error [%d]',16,1,@Status)

SET @Task = 'IF EXISTS (SELECT * FROM tempdb.dbo.sysobjects WHERE name = ' + CHAR(39) + @DBAE + CHAR(39) + ') DROP TABLE ' + @DBAE
EXECUTE (@Task)

SET NOCOUNT OFF

RETURN (@Status)


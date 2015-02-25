/*
Format same data string to html for email

*/

alter proc [dbo].[prc_mail_format_html]
@run_procedure nvarchar(max) = null,
@title nvarchar(max) = null,
@header1 nvarchar(max) = null,
@header2 nvarchar(max) = null,
@header3 nvarchar(max) = null,
@message nvarchar(max) = null,
@data_table_header nvarchar(max) = null,  
@data_table nvarchar(max) = null, -- data table in tempdb, example: 'select * from #temp'
@html nvarchar(max) = null out

as

/* TTD 
declare @html nvarchar(max)

exec [dbo].[prc_mail_format_html] 'prc.test', 'Test email', 'Header1', 'Header2', 'Header3', 
				'Same test procedure','Result table', 
				'select ''1'' as ID, ''result 1'' as Result union select ''2'', ''result 2''', 
				@html out

print @html

exec msdb.dbo.sp_send_dbmail 
			@recipients = 'avarentsov@moneks.ru', 
			@subject = 'Test Email', 
			@body_format = 'HTML',
			@body = @html
*/

declare @style nvarchar(max) = '',
		@html_table nvarchar(max) = '',	
		@server nvarchar(max) = ''

set @html = isnull(@html, '')

set @server = @@SERVERNAME

set @style = 
'
<style type="text/css">
        th {
            padding-left: 5px;
            padding-right: 5px;
            /*background: grey;
            color: white;
            border-color: white;*/
        }
        td {
            padding-left: 5px;
            padding-right: 5px;
        }
    </style>
'

set @title=
'
<html>
<head>
    <title>  
        '+isnull(@title, '')+'     
    </title>
    '+isnull(@style, '')+'
</head>
<body>
'	

--- mail header
set @html = @html + @title + '
'
set @html = @html + isnull('<h1>'+@header1+'</h1> 
', '') 
set @html = @html + isnull('<h2>'+@header2+'</h2> 
', '') 
set @html = @html + isnull('<h3>'+@header3+'</h3> 
', '') 
set @html = @html + isnull('<p>'+@message+'</p>
', '') 

-- table header
set @html = @html + isnull('<h4>'+@data_table_header+'</h4> 
', '') 

-- table 
exec [dbo].[prc_save_table_as_html] @html_table out, @DBFetch=@data_table			
set @html = @html + isnull(@html_table + '
</br>', '')

-- mail footer
set @html = @html + '<small></small>'
set @html = @html + '<footer><small> Run at ' + convert(varchar(30), getdate(), 120) + ' ' 
		+ isnull('from procedure ' + @run_procedure, '') + ' on server ' 
		+ @server + '</footer>
</body>
</html>'

go

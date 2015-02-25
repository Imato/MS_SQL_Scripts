USE [ReportServer]
GO
/****** Object:  StoredProcedure [dbo].[prc_subscription_change]    Script Date: 25.02.2015 15:48:34 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER proc [dbo].[prc_subscription_change]
@report_name varchar(255) = null,
@email_add varchar(max) = null,			
@email_remove varchar(max) = null

as
/*
exec dbo.prc_subscription_change 
		'Capture Rate by Door by Period', 
		'evladimirova@bla-bla.com;alivanova@bla-bla.com', 
		'ksarcsyan@bla-bla.com;nplekhanova@bla-bla.com;avaleeva@bla-bla.com

exec dbo.prc_subscription_change 
		'eCommerce Weekly', 
		'ostepanets@bla-bla.com', 
		'egolovanova@bla-bla.com;VLeshenko@bla-bla.com;MKavrizhnaya@bla-bla.RU;ksarcsyan@bla-bla.com'

exec dbo.prc_subscription_change 
		null, 
		'ostepanets@bla-bla.com', 
		'egolovanova@bla-bla.com;VLeshenko@bla-bla.com;MKavrizhnaya@bla-bla.RU;ksarcsyan@bla-bla.com'

*/

set nocount on

declare @version_id int

select @version_id=max(version_id)
from [dbo].[Subscriptions_bak]

if (object_id('dbo.Subscriptions_bak') is null)
begin

	create table [dbo].[Subscriptions_bak](
		[version_id] [int] NOT NULL,
		[version_date] [smalldatetime] NULL,
		[SubscriptionID] [uniqueidentifier] NOT NULL,
		[OwnerID] [uniqueidentifier] NOT NULL,
		[Report_OID] [uniqueidentifier] NOT NULL,
		[Locale] [nvarchar](128) NOT NULL,
		[InactiveFlags] [int] NOT NULL,
		[ExtensionSettings] [ntext] NULL,
		[ModifiedByID] [uniqueidentifier] NOT NULL,
		[ModifiedDate] [datetime] NOT NULL,
		[Description] [nvarchar](512) NULL,
		[LastStatus] [nvarchar](260) NULL,
		[EventType] [nvarchar](260) NOT NULL,
		[MatchData] [ntext] NULL,
		[LastRunTime] [datetime] NULL,
		[Parameters] [ntext] NULL,
		[DataSettings] [ntext] NULL,
		[DeliveryExtension] [nvarchar](260) NULL,
		[Version] [int] NOT NULL,
		[ReportZone] [int] NOT NULL
	) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

	create clustered index [cix_Subscriptions_bak_version_id] on [dbo].[Subscriptions_bak]
	(
		[version_id]
	)

end

insert into [dbo].[Subscriptions_bak]
select @version_id+1, getdate(),
	SubscriptionID, OwnerID, Report_OID, Locale, InactiveFlags, ExtensionSettings, 
	ModifiedByID, ModifiedDate, Description, LastStatus, EventType, MatchData, 
	LastRunTime, Parameters, DataSettings, DeliveryExtension, Version, ReportZone
from  [dbo].[Subscriptions]

select ArrayElem as email
into #add
from config.[dbo].[fn_textlist_to_table](@email_add, ';')

select ArrayElem as email
into #remove
from config.[dbo].[fn_textlist_to_table](@email_remove, ';')

select identity(int, 1, 1) as id, s.[SubscriptionID], s.[ExtensionSettings]
into #sub
from [dbo].[Catalog] c
join [dbo].[Subscriptions] s on c.ItemID=s.Report_OID
where c.Name like '%'+isnull(@report_name, '')+'%'

declare @i int = (select max(id) from #sub),
		@extension varchar(max),
		@extension_old varchar(max),
		@extension_new varchar(max),
		@replace_old varchar(max),
		@replace_new varchar(max),
		@emails varchar(max)

create table #emails
(email varchar(max))

while @i>0
begin

	select @extension=[ExtensionSettings],
		@extension_old=[ExtensionSettings],
		@extension_new=[ExtensionSettings]
	from #sub
	where id=@i

	-- To
	if charindex('<ParameterValue><Name>TO</Name><Value>', @extension)>0
	begin
		set @extension=substring(@extension, charindex('<ParameterValue><Name>TO</Name><Value>', @extension)+38, len(@extension))
		set @emails=substring(@extension, 1, charindex('</Value></ParameterValue>', @extension)-1)

		insert into #emails
		select ArrayElem as email
		from config.[dbo].[fn_textlist_to_table](@emails, ';')

		set @replace_old = '<ParameterValue><Name>TO</Name><Value>' + @emails + '</Value></ParameterValue>'

		--- update
		delete from #emails
		where email in (select email from #remove)

		insert into #emails
		select email from #add

		set @emails = ''
		select @emails = @emails + email + ';'
		from 
			(
			select distinct email
			from #emails
			where len(email)>5
			) t
		order by email

		set @replace_new = '<ParameterValue><Name>TO</Name><Value>' + @emails + '</Value></ParameterValue>'

		set @extension_new = replace(@extension_new, @replace_old, @replace_new)
		print 'update:' + char(13)+char(10) + @replace_old + char(13)+char(10) + ' to ' + char(13)+char(10) + @replace_new + char(13)+char(10) + ';'
	end

	delete from #emails

	-- Copy
	if charindex('<ParameterValue><Name>CC</Name><Value>', @extension)>0
	begin
		set @extension=substring(@extension, charindex('<ParameterValue><Name>CC</Name><Value>', @extension)+38, len(@extension))
		set @emails=substring(@extension, 1, charindex('</Value></ParameterValue>', @extension)-1)

		insert into #emails
		select ArrayElem as email
		from config.[dbo].[fn_textlist_to_table](@emails, ';')

		set @replace_old = '<ParameterValue><Name>CC</Name><Value>' + @emails + '</Value></ParameterValue>'

		--- update
		delete from #emails
		where email in (select email from #remove)

		set @emails = ''
		select @emails = @emails + email + ';'
		from 
			(
			select distinct email
			from #emails
			where len(email)>5
			) t
		order by email

		set @replace_new = '<ParameterValue><Name>CC</Name><Value>' + @emails + '</Value></ParameterValue>'

		set @extension_new = replace(@extension_new, @replace_old, @replace_new)
		print 'update:' + char(13)+char(10) + @replace_old + char(13)+char(10) + ' to ' + char(13)+char(10) + @replace_new + char(13)+char(10) + ';'
	end	

	delete from #emails

	update #sub
	set ExtensionSettings=@extension_new
	where id=@i
	
	set @i=@i-1

end

select s.SubscriptionID, s.ExtensionSettings as ExtensionSettings_Old,
	t.ExtensionSettings as ExtensionSettings_New
from [dbo].[Subscriptions] s
join #sub t on s.SubscriptionID=t.SubscriptionID
order by s.SubscriptionID

update s
set s.ExtensionSettings=t.ExtensionSettings
from [dbo].[Subscriptions] s
join #sub t on s.SubscriptionID=t.SubscriptionID


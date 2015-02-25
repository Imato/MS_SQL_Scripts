USE [ReportServer]
GO
/****** Object:  StoredProcedure [dbo].[prc_subscription_restore_version_last]    Script Date: 25.02.2015 15:49:38 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

ALTER proc [dbo].[prc_subscription_restore_version_last]
as
-- ReportServer.dbo.prc_subscription_restore_version_last
set nocount on

declare @version_id int

select @version_id=max([version_id])
from [dbo].[Subscriptions_bak]

update s
set s.[ExtensionSettings]=b.[ExtensionSettings]
from [dbo].[Subscriptions] s
join [dbo].[Subscriptions_bak] b
	on s.[SubscriptionID]=b.[SubscriptionID]
where b.[version_id]=@version_id

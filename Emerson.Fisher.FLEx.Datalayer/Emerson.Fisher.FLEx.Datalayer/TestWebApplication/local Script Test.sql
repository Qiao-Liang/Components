CREATE DATABASE [FlexImprovementTest]

USE [FlexImprovementTest]
GO

/****** Object:  Table [dbo].[LBPHeaderInfo]    Script Date: 10/19/2011 19:00:04 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[LBPHeaderInfo](
	[LBPID] [int] IDENTITY(0,1) NOT NULL,
	[LBPNO] [varchar](30) NOT NULL,
	[LBPName] [varchar](60) NOT NULL,
	[Address1] [varchar](100) NULL,
	[Address2] [varchar](100) NULL,
	[Address3] [varchar](100) NULL,
	[City] [varchar](60) NULL,
	[State] [varchar](60) NULL,
	[Country] [varchar](60) NULL,
	[Zip] [varchar](20) NULL,
	[PhoneNo] [varchar](30) NULL,
	[ContactName] [varchar](60) NULL,
 CONSTRAINT [PK_LBP] PRIMARY KEY CLUSTERED 
(
	[LBPID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO

INSERT INTO [FlexImprovementTest].[dbo].[LBPHeaderInfo] 
(
[LBPNO],[LBPName],[Address1],[Address2],[Address3],[City],[State],
[Country],[Zip],[PhoneNo],[ContactName]) 
VALUES
('1','test1','test1','test1','test1','test1','test1','test1','test1','test1','test1')
GO

USE [FlexImprovementTest]
GO

/****** Object:  StoredProcedure [dbo].[deletedate]    Script Date: 10/19/2011 19:02:09 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[deletedate]
	-- Add the parameters for the stored procedure here
	@para int
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;
delete  FROM [FlexImprovementTest].[dbo].[LBPHeaderInfo]
where LBPID=@para

END

GO
USE [FlexImprovementTest]
GO

/****** Object:  StoredProcedure [dbo].[InserandReturnPrimay]    Script Date: 10/19/2011 19:02:43 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[InserandReturnPrimay]
	-- Add the parameters for the stored procedure here
@para int 
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
	INSERT INTO [FlexImprovementTest].[dbo].[LBPHeaderInfo] 
(
[LBPNO],[LBPName],[Address1],[Address2],[Address3],[City],[State],
[Country],[Zip],[PhoneNo],[ContactName]) 
VALUES
('3','test1','test1','test1','test1','test1','test1','test1','test1','test1','test1')

select top 1 LBPID from [FlexImprovementTest].[dbo].[LBPHeaderInfo] order by LBPID desc

END

GO

USE [FlexImprovementTest]
GO

/****** Object:  StoredProcedure [dbo].[InsertDATA]    Script Date: 10/19/2011 19:03:09 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[InsertDATA]
	-- Add the parameters for the stored procedure here
	@PARA int
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;
INSERT INTO [FlexImprovementTest].[dbo].[LBPHeaderInfo] 
(
[LBPNO],[LBPName],[Address1],[Address2],[Address3],[City],[State],
[Country],[Zip],[PhoneNo],[ContactName]) 
VALUES
('3','test1','test1','test1','test1','test1','test1','test1','test1','test1','test1')
	
END

GO


USE [FlexImprovementTest]
GO

/****** Object:  StoredProcedure [dbo].[MutipleReturn]    Script Date: 10/19/2011 19:03:44 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[MutipleReturn]
	-- Add the parameters for the stored procedure here
	@para1 int 
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SELECT TOP 10 [LBPID]
      ,[LBPNO]
      ,[LBPName]
      ,[Address1]
      ,[Address2]
      ,[Address3]
      ,[City]
      ,[State]
      ,[Country]
      ,[Zip]
      ,[PhoneNo]
      ,[ContactName]
  FROM [FlexImprovementTest].[dbo].[LBPHeaderInfo] order by [LBPID] desc
  
  select [LBPID] FROM [FlexImprovementTest].[dbo].[LBPHeaderInfo] order by [LBPID] desc

END

GO


USE [FlexImprovementTest]
GO

/****** Object:  StoredProcedure [dbo].[updatedate]    Script Date: 10/19/2011 19:04:02 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[updatedate]
	-- Add the parameters for the stored procedure here
@pare int 	
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
	UPDATE [FlexImprovementTest].[dbo].[LBPHeaderInfo]
   SET 
      [LBPName] = 'rep1'
      ,[Address1] = 'rep1'
      ,[Address2] = 'rep1'
      ,[Address3] = 'rep1'
      ,[City] = 'rep1'
      ,[State] ='rep1'
      ,[Country] = 'rep1'
      ,[Zip] = 'rep1'
      ,[PhoneNo] = 'rep1'
      ,[ContactName] = 'rep1'
 WHERE LBPID=@pare
END

GO
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
begin tran

CREATE TABLE [dbo].[LOCAL_OE_ORDER_HEADERS_ALL](
	ORDER_NUMBER [int] NOT NULL,
	CUST_PO_NUMBER [varchar](1000) NOT NULL,
	SHIPPING_METHOD_CODE [varchar](1000) NOT NULL,
	PACKING_INSTRUCTIONS [varchar](1000) NULL,
)

insert into [LOCAL_OE_ORDER_HEADERS_ALL] values (100000,'SSP059292','000001_UPS_T_GND','Attention To: Sebastian equipment Company; PO # SSP059292')
insert into [LOCAL_OE_ORDER_HEADERS_ALL] values (100002,'863 -TJM43751','000001_UPS_T_GND','C/R; 863 - TJM4351A/ Item NO:2/Tag No: FV-4036-6; FV-4037-7')
insert into [LOCAL_OE_ORDER_HEADERS_ALL] values (100006,'B049263776','000001_UPS_T_GND','B032227925')
insert into [LOCAL_OE_ORDER_HEADERS_ALL] values (100007,'107-D501221','000001_DHL_A_XPR','S/O# 100007  107-D501221  22710486')
insert into [LOCAL_OE_ORDER_HEADERS_ALL] values (100008,'W004938','000001_UPS BLUE_A_1DA','Attn: Goodrich')
insert into [LOCAL_OE_ORDER_HEADERS_ALL] values (100011,'122887','000001_UPS_T_GND','PO# 122887')
insert into [LOCAL_OE_ORDER_HEADERS_ALL] values (100013,'014M-B206167497','000001_UPS_T_GND','945282')
insert into [LOCAL_OE_ORDER_HEADERS_ALL] values (100015,'2341-11339','000001_UPS_T_GND','P.O. 2341-11339')
insert into [LOCAL_OE_ORDER_HEADERS_ALL] values (100039,'099X-16421','000001_UPS RED_A_1DA','Attention Bob Riley/Fisher Education Services')
insert into [LOCAL_OE_ORDER_HEADERS_ALL] values (100018,'124I-119003','000001_DHL_A_XPR','117228')

rollback tran




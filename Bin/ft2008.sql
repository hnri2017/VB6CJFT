--===========================================================================
--����˵����
--1����SQLServer2008���ù���ԱȨ��(����ͬ)ִ�����´��롣
--2������˳����ò�Ҫ�ҡ����빦���ǽ������ݿ⡢���洢���̵���Ϣ��
--3������һ���ܽ�����ɾ�Ĳ����ݵ��˺Ÿ������ݿ�(����Ҫ�½������ݿ�����˺�)��
--4���������ݿ�IP�����ݿ����������˺����������õ�����˳����С�
--===========================================================================

--==================�ָ���===================================================

--===========================================================================
--�������������ݿ�
--===========================================================================
USE [master]
GO

/****** Object:  Database [db_FT]    Script Date: 09/18/2018 08:39:13 ******/
CREATE DATABASE [db_FT] ON  PRIMARY 
( NAME = N'db_FT', FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL10.MSSQLSERVER\MSSQL\DATA\db_FT.mdf' , SIZE = 3072KB , MAXSIZE = UNLIMITED, FILEGROWTH = 1024KB )
 LOG ON 
( NAME = N'db_FT_log', FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL10.MSSQLSERVER\MSSQL\DATA\db_FT_log.ldf' , SIZE = 1024KB , MAXSIZE = 2048GB , FILEGROWTH = 10%)
GO

ALTER DATABASE [db_FT] SET COMPATIBILITY_LEVEL = 100
GO

IF (1 = FULLTEXTSERVICEPROPERTY('IsFullTextInstalled'))
begin
EXEC [db_FT].[dbo].[sp_fulltext_database] @action = 'enable'
end
GO

ALTER DATABASE [db_FT] SET ANSI_NULL_DEFAULT OFF 
GO

ALTER DATABASE [db_FT] SET ANSI_NULLS OFF 
GO

ALTER DATABASE [db_FT] SET ANSI_PADDING OFF 
GO

ALTER DATABASE [db_FT] SET ANSI_WARNINGS OFF 
GO

ALTER DATABASE [db_FT] SET ARITHABORT OFF 
GO

ALTER DATABASE [db_FT] SET AUTO_CLOSE OFF 
GO

ALTER DATABASE [db_FT] SET AUTO_CREATE_STATISTICS ON 
GO

ALTER DATABASE [db_FT] SET AUTO_SHRINK OFF 
GO

ALTER DATABASE [db_FT] SET AUTO_UPDATE_STATISTICS ON 
GO

ALTER DATABASE [db_FT] SET CURSOR_CLOSE_ON_COMMIT OFF 
GO

ALTER DATABASE [db_FT] SET CURSOR_DEFAULT  GLOBAL 
GO

ALTER DATABASE [db_FT] SET CONCAT_NULL_YIELDS_NULL OFF 
GO

ALTER DATABASE [db_FT] SET NUMERIC_ROUNDABORT OFF 
GO

ALTER DATABASE [db_FT] SET QUOTED_IDENTIFIER OFF 
GO

ALTER DATABASE [db_FT] SET RECURSIVE_TRIGGERS OFF 
GO

ALTER DATABASE [db_FT] SET  DISABLE_BROKER 
GO

ALTER DATABASE [db_FT] SET AUTO_UPDATE_STATISTICS_ASYNC OFF 
GO

ALTER DATABASE [db_FT] SET DATE_CORRELATION_OPTIMIZATION OFF 
GO

ALTER DATABASE [db_FT] SET TRUSTWORTHY OFF 
GO

ALTER DATABASE [db_FT] SET ALLOW_SNAPSHOT_ISOLATION OFF 
GO

ALTER DATABASE [db_FT] SET PARAMETERIZATION SIMPLE 
GO

ALTER DATABASE [db_FT] SET READ_COMMITTED_SNAPSHOT OFF 
GO

ALTER DATABASE [db_FT] SET HONOR_BROKER_PRIORITY OFF 
GO

ALTER DATABASE [db_FT] SET  READ_WRITE 
GO

ALTER DATABASE [db_FT] SET RECOVERY FULL 
GO

ALTER DATABASE [db_FT] SET  MULTI_USER 
GO

ALTER DATABASE [db_FT] SET PAGE_VERIFY CHECKSUM  
GO

ALTER DATABASE [db_FT] SET DB_CHAINING OFF 
GO


--==================�ָ���===================================================

--===========================================================================
--������������[tb_FT_Sys_User]�������˺�����
--===========================================================================
USE [db_FT]
GO

/****** Object:  Table [dbo].[tb_Test_Sys_User]    Script Date: 2018/9/15 21:48:49 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[tb_FT_Sys_User](
	[UserAutoID] [int] IDENTITY(2000,1) NOT NULL,
	[UserLoginName] [nvarchar](50) NOT NULL,
	[UserPassword] [nvarchar](50) NOT NULL,
	[UserFullName] [nvarchar](50) NULL,
	[UserSex] [nvarchar](2) NULL,
	[UserState] [nvarchar](50) NULL,
	[DeptID] [int] NULL,
	[UserMemo] [nvarchar](500) NULL,
 CONSTRAINT [PK_tb_FT_Sys_User] PRIMARY KEY CLUSTERED 
(
	[UserAutoID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
--===========================================================================
--��������[tb_FT_Sys_User]���в���Ĭ�ϵĳ���Ȩ���˺�����
--===========================================================================
USE [db_FT]
GO

INSERT INTO [db_FT].[dbo].[tb_FT_Sys_User]([UserLoginName],[UserPassword],[UserFullName])
VALUES('admin','9E7445656E63AB22FC3EA4387D00','��������Ա')	--����ftadmin
GO

INSERT INTO [db_FT].[dbo].[tb_FT_Sys_User]([UserLoginName],[UserPassword],[UserFullName])
VALUES('system','9E7445657C63B622E23EB93876000744','ϵͳ����Ա')	--����ftsystem
GO

--==================�ָ���===================================================

--===========================================================================
--�����������洢����[sp_FT_Sys_UserLogin]
--===========================================================================
USE [db_FT]
GO

/****** Object:  StoredProcedure [dbo].[sp_FT_Sys_UserLogin]    Script Date: 2018/9/15 23:27:51 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[sp_FT_Sys_UserLogin] 
	-- Add the parameters for the stored procedure here
	@strUN AS NVARCHAR(50)
	,@strPWD AS NVARCHAR(50)
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
	SELECT * 
	FROM tb_FT_Sys_User 
	WHERE UserLoginName = @strUN AND UserPassword =@strPWD 

END

GO

--===========================================================================
--������
--===========================================================================


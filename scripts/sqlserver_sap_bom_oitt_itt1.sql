-- SAP B1 BOM staging (same column layout as OITT_TMP.xlsx / ITT1_TMP.xlsx).
-- Database: SW (or change USE). Run in SSMS after selecting DB SW.
-- Tables: dbo.OITT_TMP, dbo.ITT1_TMP

IF OBJECT_ID(N'dbo.OITT_TMP', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.OITT_TMP (
        [#]                              INT            NULL,
        [Parent Item]                    NVARCHAR(64)   NULL,
        [BOM Type]                       NVARCHAR(8)    NULL,
        [Price List]                     INT            NULL,
        [No. of Units]                   INT            NULL,
        [Creation Date]                  NVARCHAR(32)   NULL,
        [Date of Update]                 NVARCHAR(32)   NULL,
        [Postponed to Next Year]         NVARCHAR(8)    NULL,
        [Data source]                    NVARCHAR(8)    NULL,
        [User Signature]                 INT            NULL,
        [SCN Counter]                    INT            NULL,
        [Display Currency]               INT            NULL,
        [Whse for Finished Product]      NVARCHAR(16)   NULL,
        [Object Type]                    NVARCHAR(16)   NULL,
        [Log Instance - History]         INT            NULL,
        [Updating User]                  INT            NULL,
        [Distribution Rule]              NVARCHAR(128)  NULL,
        [Hide Components in Printing]    NVARCHAR(8)    NULL,
        [Distribution Rule2]             NVARCHAR(128)  NULL,
        [Distribution Rule3]             NVARCHAR(128)  NULL,
        [Distribution Rule4]             NVARCHAR(128)  NULL,
        [Distribution Rule5]             NVARCHAR(128)  NULL,
        [Time of Update]                 NVARCHAR(32)   NULL,
        [Project Code]                   NVARCHAR(64)   NULL,
        [Planned Average Production Size] INT           NULL,
        [Product Description]            NVARCHAR(512)  NULL,
        [Create Time - Incl. Secs]       BIGINT         NULL,
        [Update Full Time]               BIGINT         NULL,
        [Attachment Entry]               INT            NULL,
        [Attachments]                    INT            NULL
    );
END
GO

-- To replace dbo.ITT1_TMP entirely (drop all columns / old layout), run:
--   DROP TABLE IF EXISTS dbo.ITT1_TMP;
-- then execute the CREATE TABLE dbo.ITT1_TMP block below (not only the IF NOT EXISTS branch).

IF OBJECT_ID(N'dbo.ITT1_TMP', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ITT1_TMP (
        [#]                          INT             NULL,
        [Parent Item]                NVARCHAR(64)    NULL,
        [Component Element Number]   INT             NULL,
        [Visual Order]               INT             NULL,
        [Component Code]           NVARCHAR(64)    NULL,
        [Quantity]                   DECIMAL(18, 6)  NULL,
        [Warehouse]                  NVARCHAR(16)    NULL,
        [Price]                      DECIMAL(18, 6)  NULL,
        [Currency]                   NVARCHAR(16)    NULL,
        [Price List]                 INT             NULL,
        [Original Price]             DECIMAL(18, 6)  NULL,
        [Original Currency]          NVARCHAR(16)    NULL,
        [Issue Method]               NVARCHAR(8)     NULL,
        [Inventory UoM]              NVARCHAR(32)    NULL,
        [Comment]                    NVARCHAR(512)   NULL,
        [Log Instance]               INT             NULL,
        [Object]                     NVARCHAR(16)    NULL,
        [Distribution Rule]          NVARCHAR(128)   NULL,
        [Distribution Rule2]         NVARCHAR(128)   NULL,
        [Distribution Rule3]         NVARCHAR(128)   NULL,
        [Distribution Rule4]         NVARCHAR(128)   NULL,
        [Distribution Rule5]         NVARCHAR(128)   NULL,
        [Principal Input]            NVARCHAR(8)     NULL,
        [Project Code]               NVARCHAR(64)    NULL,
        [Component Type]             INT             NULL,
        [WIP Account Code]           NVARCHAR(64)    NULL,
        [Additional Quantity]        DECIMAL(18, 6)  NULL,
        [Row Text]                   NVARCHAR(512)   NULL,
        [Stage ID]                   NVARCHAR(64)    NULL,
        [Item Description]           NVARCHAR(512)   NULL
    );
END
GO

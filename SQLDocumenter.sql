USE tempdb
GO
IF CAST(SERVERPROPERTY('productversion') as varchar(2)) = '8.'
RAISERROR ('Version of SQL Server is not supported', 20, -1) with log;
GO
DECLARE 
	@DBName VARCHAR(128), 
	@Collation VARCHAR(50), 
	@ReportOnlyObjectNames TINYINT,			
	@CheckFunctionDependents TINYINT,		
	@Debug TINYINT;
SELECT @DBName = '<YOUR DATABASE NAME>',
	@Collation = 'SQL_Latin1_General_CP1_CI_AS',
	@ReportOnlyObjectNames = 0,			/* 0/1 - That function allows to read only object names (for databases with thousands of objects) */
	@CheckFunctionDependents = 0,			/* 0/1 - Check for Function dependencies within Stored Procedures, Triggers and other functions */
	@Debug = 0;

/********************************************************************************************************************************************************************************
#Description: Script Generates full documentation for a database
#Details:
#Contacts:
- Twitter: @SlavaSQL
- Blog: http://slavasql.blogspot.com/2013/11/stored-procedure-to-document-database.html

#Permissions: Requires at least to be a member of db_datareader role within the documented database 
  To perform full capacity documenting need to have also following permissions:
- VIEW DATABASE STATE (to see sizes of objects and list all tables)
- VIEW DEFINITION (To see all programmable objects)
- VIEW SERVER STATE (to see size of log file)
- To get Number of VLFs in the log file permissions have to be escalated to sysadmin role 

#Tested on:
- 2016 Enterprise
- 2014 Enterprise
- 2012 Enterprise,Standard.
- 2008 R2 (SP1) Enterprise.
- 2005 Express.

#Execution:
- Procedure has to be executed under full administrative privileges
- Results of the procedure generate browsable HTML document

Script does not cover:
- function_order_columns;
- XML Schema types;
- FS = Assembly (CLR) scalar-function
- FT = Assembly (CLR) table-valued function
- IF = SQL inline table-valued function
- IT = Internal table
- PC = Assembly (CLR) stored-procedure
- PG = Plan guide
- R = Rule (old-style, stand-alone)
- RF = Replication-filter-procedure
- S = System base table
- SQ = Service queue
- TT = Table type
- TA = Assembly (CLR) DML trigger
- X = Extended stored procedure

# Procedure has not been tested with (yet):
- Columnstore indexes;
- Encrypted Code;
- filestream files/tables;
- Busy OLTP server;
- (a lot of other functionality)

# In Line:
* Index usage;
- Show that table/index is partitioned
- Search for synonym's usage.
- Hide/show lists of elements
- Add progress reporting for lists of objects bigger than 1000
- Add partitioning to statistics
- Table valued function columns
- Table valued function constraints
- Ext Properties of Functions' and Procedures' parameters. 
- Synonym dependencies
- Usage of Objects from other databases.

BUG: List of Partitioned Objects has a lot of nulls.
* Move "FileGroup Name" before "Object Schema"

BUGs: 
- Empty Lists of Partition functions and partition schemas 
- Extended properties for parameters of stored procedures and functions are not visible 

# Plans:
- user accounts;
- Spatial Indexes;
- run against linked server.

# Thanks to
Pinal Dave
Olaf Helper

#Parameters
1. @DBName (required) - database name
2. @CheckFunctionDependents - specify "1" if want to check possible function dependency (works slower)
3. @Debug - specify "1" if want to see all intermediate queries

#Example
-- Real results
DECLARE @i INT
EXEC @i = usp_Documenting_DB @DBName='TempDB'
PRINT CAST(@i AS VARCHAR)
-- Test
DECLARE @i INT
EXEC @i = usp_Documenting_DB @DBName='AdventureWorks', @Debug=1
PRINT CAST(@i AS VARCHAR)

#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#Cahnge History:
Vers. | Operator            |   Date   |  Action     | Description
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
01.00 | Slava Murygin       |2013-11-11| Creation    | Initial Creation of the Stored Procedure
01.01 | Slava Murygin       |2013-12-06| Fix Bug     | In 2008 and earlier versions  Free space in data file was calculated incorrectly (Q:0030). 
01.02 | Slava Murygin       |2014-01-05| Enhancement | Handle Multi-line Extended Properties (Q:0015). Added DB Schemas, Filegroups. Added showing Extended properties to all objects.
01.03 | Slava Murygin       |2014-01-06| Enhancement | Show: Tables' triggers. Triggers' properties, Database Triggers. FK Reference. Table dependancy via FK. 
01.04 | Slava Murygin       |2014-01-07| Enhancement | Added Sever version Recognition. Show: Partition(ed) Functions, Objects, Schemas.
01.05 | Slava Murygin       |2014-01-21| Enhancement | Added Error Handling. Changed to Simple script format
01.06 | Slava Murygin       |2014-01-28| Fix Bug     | Fix incorrect presenting of DB extended properties and Server version. Added collation issues handling.
01.07 | Slava Murygin       |2014-02-05| Fix Bug     | Allow script to run under server collation "SQL_AltDiction_Pref_CP850_CI_AS" + added immediate step completion.
01.08 | Slava Murygin       |2014-02-05| Enhancement | Added availability to document only lists of objects (for DBs with thouthands of objects)
01.09 | Slava Murygin       |2014-02-12| Enhancement | Added: hiding for object lists, Name of File Groups in the list of List of Partition Functions, table sizes.
01.10 | Slava Murygin       |2014-02-12| Fix Bug     | Handled Null Value in list of Partitioned objects. Exclude Diagram-Extended properies for views.
01.11 | Slava Murygin       |2014-02-13| Enhancement | Added: Tables' Statistics Info; Show schema for "dependency" objects; Hiding on Objects' Details.
01.12 | Slava Murygin       |2014-02-13| Enhancement | Added: DB Backups; # of VLFs;
01.13 | Slava Murygin       |2014-02-15| Fix Bug     | Handling of LogInf for V.2014; Not showing empty Patrition Schemas/Functions tables;
01.14 | Slava Murygin       |2014-02-17| Fix Bug     | Number of records in a table;
01.15 | Slava Murygin       |2014-05-13| Fix Bug     | DBCC LogInfo has 8 columns after SS2012;
01.16 | Slava Murygin       |2014-05-16| Enhancement | Implemented Percentage tracker when @ReportOnlyObjectNames = 0 ;
01.17 | Slava Murygin       |2014-05-17| Enhancement | Tuned some queries for better performance ;
01.18 | Slava Murygin       |2014-05-21| Enhancement | Better handling of permission errors ;
01.19 | Slava Murygin       |2014-05-21| Enhancement | DB Options; Filegroups to DB files' table; DB Logins' List; Tables' Comression;
01.20 | Slava Murygin       |2014-05-21| Fix Bug     | Filter cells in "Statistics" table was not handled properly; 
01.21 | Slava Murygin       |2014-05-21| Fix Bug     | List of dependencies printed even if they are empty
01.22 | Slava Murygin       |2014-05-25| Downgrading | Refubrished script to meet SQL 2005 criterias
01.23 | Slava Murygin       |2018-01-28| Enhancement | Added support of SQL Server 2016
01.24 | Slava Murygin       |2022-05-22| Enhancement | Added support of SQL Server 2017-19
#*******************************************************************************************************************************************************************************/

SET NOCOUNT ON

DECLARE @Results TABLE (ID INT IDENTITY(1,1), ResultRecord VARCHAR(MAX)); /* Table to collect ALL results*/
DECLARE @IntermediateResults VARCHAR(MAX); /* Used to store intermediate results before inserting into @Results table */ 
DECLARE @TempString NVARCHAR(MAX); /* Temporary String Variable. Mostly Used for SQL Queries */
DECLARE @Objects TABLE (
       ID INT IDENTITY(1,1), 
       SchemaName VARCHAR(128), 
       SchemaID  INT, 
       ObjectName VARCHAR(128),  
       ObjectID  INT, 
       ParentObjectID INT, 
       Created_Dt DATETIME, 
       Modified_Dt DATETIME, 
       ObjectType CHAR(2), 
       ObjectType_Desc VARCHAR(60),
       Ext_Property VARCHAR(MAX),
       Reported TINYINT DEFAULT (0)
       ) ;
DECLARE @object_id  INT; /* Stores current Object ID */
DECLARE @SessionId CHAR(36)  ; /* Unique ID for all temp objects within the SP*/
DECLARE @i INT; /* Usualy Counter */ 
DECLARE @TotalNumber INT, @CurrentNumber INT, @IncrementCount INT, @CntMessage VARCHAR(100); /* Set of Counters for percentage indicator */
DECLARE @10Percent INT, @5Percent INT, @NextMilestone INT; /* Percentage thresholds. if bigger than @5Percent then uses 1% threshold */
DECLARE @v TINYINT; /* UIndicate possible violations */ 
DECLARE @vt TABLE (ViolationID TINYINT, Violated INT DEFAULT 0);
DECLARE @OType_Desc nvarchar(128); /* Store Object Type Description */
DECLARE @OType char(2); /* Store Object Type */
DECLARE @ObjStats TABLE(StatID INT IDENTITY(1,1), ObjectTypeDesc NVARCHAR(60), ObjectType CHAR(2), ObjCnt INT); /* Object statistics */
DECLARE @Dependent TABLE(ID INT IDENTITY(1,1), [Object_id] INT NOT NULL, [Schema_Name] VARCHAR(128), [Object_Name] VARCHAR(128), ObjectTypeDesc VARCHAR(60)); /* Used to determine functions' depent objects */
DECLARE @LogInfo TABLE (RecoveryUnitId BIGINT, FileId INT, FileSize BIGINT, StartOffset BIGINT, FSeqNo BIGINT, Status BIGINT, Parity BIGINT, CreateLSN NUMERIC(38), ID INT IDENTITY(1,1)); /* Used to calculate number of VLFs */
DECLARE @Obj_Name VARCHAR(128); /* Holds Object name*/
DECLARE @ServerVersion VARCHAR(15); /* Version of SQL Server this SP is Running on */
DECLARE @ServerRelease SMALLINT; /* Version of SQL Server this SP is Running on */
DECLARE @ErrorList VARCHAR(MAX); /* List of Errors during an Execution */
DECLARE @ErrorMessage VARCHAR(2048); /* Temporary Error Variable */
DECLARE @CurrentStep CHAR(4);

/* Set Defaults (Required by 2005)*/
SELECT @i = 0,
	@SessionId = REPLACE(CAST(NewID() AS CHAR(36)),'-','_'),
	@ErrorList = '',
	@TotalNumber = 0,
	@CurrentNumber = 0,
	@IncrementCount = 0,
	@10Percent = 1000,
	@5Percent = 10000,
	@NextMilestone = 0,
	@v = 0;

/*0010 Verify if Database Exists */
SET @CurrentStep = '0010';
SELECT TOP 1 @i = database_id FROM master.sys.databases WHERE @DBName = name
IF IsNull(@i,0) = 0
BEGIN
  RAISERROR('Database does not Exist',16,1)
  GOTO END_of_SCRIPT
END

RAISERROR ('#0010 Finished', 0, 1) WITH NOWAIT;

/*0012 Verify if SQL Server Version */
SET @CurrentStep = '0012';
SELECT @ServerVersion = CAST(SERVERPROPERTY('productversion') as VARCHAR(15)), 
  @ServerRelease = CAST(LEFT(@ServerVersion, PATINDEX('%.%', @ServerVersion)-1) as SMALLINT) * 10,
  @ServerVersion = 
  CASE @ServerRelease
    WHEN 90 THEN '2005'
    WHEN 100 THEN '2008'
    WHEN 110 THEN '2012'
    WHEN 120 THEN '2014'
    WHEN 130 THEN '2016'
    WHEN 140 THEN '2017'
    WHEN 150 THEN '2019'
    ELSE CAST(@ServerRelease AS VARCHAR)
  END

PRINT 'SQL Server ' + @ServerVersion

IF @ServerVersion is Null
BEGIN
  RAISERROR('This Version of SQL Server is not Supported',16,1)
  GOTO END_of_SCRIPT
END

RAISERROR ('#0012 Finished', 0, 1) WITH NOWAIT;

/*0014 Insert Dummy Error Record */
SET @CurrentStep = '0014';
INSERT INTO @Results(ResultRecord) VALUES(@ErrorList);

RAISERROR ('#0014 Finished', 0, 1) WITH NOWAIT;

/*0015 Collecting Extended Properties */
SET @CurrentStep = '0015';

/* Create temporary table */
SET @TempString = 'CREATE TABLE ##Temp_extended_Properties_' + @SessionId + 
  '(class tinyint, major_id int, minor_id int, class_desc varchar(25), value VARCHAR(MAX), PRIMARY KEY (class,major_id,minor_id))'
IF @Debug = 1 PRINT @TempString;

  BEGIN TRY
    EXECUTE (@TempString);
  END TRY

  BEGIN CATCH
    PRINT @TempString; PRINT ERROR_MESSAGE();
    RAISERROR (@ErrorMessage, 16, 0);
    GOTO END_of_SCRIPT
  END CATCH


/* Collect Extended Properties into temporary table */
SET @TempString = '
  ;WITH d AS (SELECT class, major_id, minor_id, class_desc FROM ' + @DBName + '.sys.extended_properties 
    GROUP BY class, major_id, minor_id, class_desc HAVING COUNT(*) > 1)
  INSERT INTO ##Temp_extended_Properties_' + @SessionId + '(class, major_id, minor_id, class_desc, value)
  SELECT d.class, d.major_id, d.minor_id, CAST(d.class_desc as varchar(25)), 
    REPLACE(
      SUBSTRING (( 
        SELECT CHAR(255) + CASE Name WHEN ''MS_Description'' THEN '''' ELSE Name + N'': '' END + CAST(Value as VARCHAR(MAX)) 
        FROM ' + @DBName + '.sys.extended_properties as e WHERE e.major_id = d.major_id and e.minor_id = d.minor_id and e.class = d.class 
			and Name not like ''MS_DiagramPane%''
        FOR XML PATH ('''')
      ),2, 8000  ), CHAR(255), ''<BR>'')
  FROM d
  UNION
  SELECT ep.class, ep.major_id, ep.minor_id, CAST(ep.class_desc as varchar(25)), 
    CASE Name WHEN ''MS_Description'' THEN '''' ELSE Name + N'': '' END + CAST(Value as VARCHAR(MAX)) COLLATE ' + @Collation + '
  FROM ' + @DBName + '.sys.extended_properties as ep
  LEFT JOIN d ON ep.major_id = d.major_id and ep.minor_id = d.minor_id and ep.class = d.class 
  WHERE d.class is Null;';
IF @Debug = 1 PRINT @TempString;

  BEGIN TRY
    EXECUTE (@TempString);
  END TRY

  BEGIN CATCH
    SET @ErrorMessage = ERROR_MESSAGE();
        
    BEGIN TRY
        RAISERROR (@ErrorMessage, 16, 0);
    END TRY
    
    BEGIN CATCH
      PRINT @TempString; PRINT ERROR_MESSAGE();
      SELECT @ErrorList = @ErrorList + 
        'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
        'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
    END CATCH
  END CATCH

RAISERROR ('#0015 Finished', 0, 1) WITH NOWAIT;

/*0020 Collecting General Server Information */
SET @CurrentStep = '0020';
INSERT INTO @Results(ResultRecord) VALUES 
  ('<html><head></head>');
INSERT INTO @Results(ResultRecord) VALUES 
  ('<script language=javascript>');
INSERT INTO @Results(ResultRecord) VALUES 
  ('function HideShowCode(hs,ObjectHS){if (hs==0) {document.getElementById(ObjectHS).style.display="none";} else {document.getElementById(ObjectHS).style.display="block";}}');
INSERT INTO @Results(ResultRecord) VALUES 
  ('</script>');
INSERT INTO @Results(ResultRecord) VALUES 
  ('<body><Center>' +
  '<H1>Full Database Documentation</H1>' +
  '<H2>For Database "' + @DBName + '" (Database ID = ' + CAST(@i as VARCHAR) + ' )</H2>');

SET @TempString = '
  SELECT TOP 1 ''<TABLE><TR><TD VALIGN="TOP"><H3>DB Description: </H3></TD><TD><H4>'' + value + ''</H4></TD></TR></TABLE>''
  FROM ##Temp_extended_Properties_' + @SessionId + '
  WHERE class = 0;';
IF @Debug = 1 PRINT @TempString;

  BEGIN TRY
    INSERT INTO @Results(ResultRecord) EXECUTE (@TempString);
  END TRY

  BEGIN CATCH
    SET @ErrorMessage = ERROR_MESSAGE();
    
    BEGIN TRY
        RAISERROR (@ErrorMessage, 16, 0);
    END TRY
    
    BEGIN CATCH
      PRINT @TempString; PRINT ERROR_MESSAGE();
      SELECT @ErrorList = @ErrorList + 
        'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
        'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
    END CATCH
  END CATCH
  
INSERT INTO @Results(ResultRecord)
VALUES 
  ('</Center><H3>SQL server Name "' + @@Servername + '"</H3>' + 
  '<B>SQL Server Version:</B><BR/><pre>' + REPLACE(@@VERSION,CHAR(10),'<BR>') + '</pre>' +
  '<H2>Database Highlights:</H2>')

RAISERROR ('#0020 Finished', 0, 1) WITH NOWAIT;

/*0022 Collecting DB Options Information */
SET @CurrentStep = '0022';
INSERT INTO @Results(ResultRecord)
VALUES ('<table border=1 cellpadding=5 style="font-weight:bold;"><TR><TH>Database Option</TH><TH>Option''s Value</TH><TH>Database Option</TH><TH>Option''s Value</TH></TR>' ) ;

SET @TempString = '
SELECT 
''<TR>'' + 
''<TD>Database ID</TD><TD ALIGN="CENTER">'' + CAST(database_id as VARCHAR) + ''</TD>'' + 
''<TD>Creation Date</TD><TD>'' + CONVERT(VARCHAR,create_date,120) + ''</TD>'' + 
''</TR>'' + 

''<TR>'' + 
''<TD>READ_ONLY</TD><TD ALIGN="CENTER">'' + CASE is_read_only WHEN 0 THEN ''<FONT COLOR="BLUE">OFF</FONT>'' ELSE ''<FONT COLOR="GREEN">ON</FONT>'' END + ''</TD>'' + 
''<TD>DB State</TD><TD>'' + state_desc + ''</TD>'' + 
''</TR>'' + 

''<TR>'' + 
''<TD>DB_ENCRYPTED</TD><TD ALIGN="CENTER">'' + CASE ' + CASE WHEN @ServerRelease >= 100 THEN 'is_encrypted' ELSE '0' END +
				' WHEN 0 THEN ''<FONT COLOR="BLUE">OFF</FONT>'' ELSE ''<FONT COLOR="GREEN">ON</FONT>'' END + ''</TD>'' + 
''<TD>User Access</TD><TD>'' + user_access_desc COLLATE SQL_Latin1_General_CP1_CI_AS + ''</TD>'' + 
''</TR>'' + 

''<TR>'' + 
''<TD>AUTO_SHRINK</TD><TD ALIGN="CENTER">'' + CASE is_auto_shrink_on WHEN 0 THEN ''<FONT COLOR="BLUE">OFF</FONT>'' ELSE ''<FONT COLOR="GREEN">ON</FONT>'' END + ''</TD>'' + 
''<TD>DB recovery model</TD><TD>'' + recovery_model_desc + ''</TD>'' + 
''</TR>'' + 
 
''<TR>'' + 
''<TD>DB IN Standby</TD><TD ALIGN="CENTER">'' + CASE is_in_standby WHEN 0 THEN ''<FONT COLOR="BLUE">OFF</FONT>'' ELSE ''<FONT COLOR="GREEN">ON</FONT>'' END + ''</TD>'' + 
''<TD>DB Collation</TD><TD>'' + collation_name + ''</TD>'' + 
''</TR>'' + 

''<TR>'' + 
''<TD>ANSI_NULLS</TD><TD ALIGN="CENTER">'' + CASE is_ansi_nulls_on WHEN 0 THEN ''<FONT COLOR="BLUE">OFF</FONT>'' ELSE ''<FONT COLOR="GREEN">ON</FONT>'' END + ''</TD>'' + 
''<TD>AUTO_CREATE_STATISTICS</TD><TD>'' + CASE is_auto_create_stats_on WHEN 0 THEN ''<FONT COLOR="BLUE">OFF</FONT>'' ELSE ''<FONT COLOR="GREEN">ON</FONT>'' END + ''</TD>'' + 
''</TR>'' + 

''<TR>'' + 
''<TD>ANSI_NULL_DEFAULT</TD><TD ALIGN="CENTER">'' + CASE is_ansi_null_default_on WHEN 0 THEN ''<FONT COLOR="BLUE">OFF</FONT>'' ELSE ''<FONT COLOR="GREEN">ON</FONT>'' END + ''</TD>'' + 
''<TD>AUTO_UPDATE_STATISTICS</TD><TD>'' + CASE is_auto_update_stats_on WHEN 0 THEN ''<FONT COLOR="BLUE">OFF</FONT>'' ELSE ''<FONT COLOR="GREEN">ON</FONT>'' END + ''</TD>'' + 
''</TR>'' + 
''<TR>'' + 
''<TD>ANSI_WARNINGS</TD><TD ALIGN="CENTER">'' + CASE is_ansi_warnings_on WHEN 0 THEN ''<FONT COLOR="BLUE">OFF</FONT>'' ELSE ''<FONT COLOR="GREEN">ON</FONT>'' END + ''</TD>'' + 
''<TD>AUTO_UPDATE_STATISTICS_ASYNC</TD><TD>'' + CASE is_auto_update_stats_async_on WHEN 0 THEN ''<FONT COLOR="BLUE">OFF</FONT>'' ELSE ''<FONT COLOR="GREEN">ON</FONT>'' END + ''</TD>'' + 
''</TR>'' + 

''<TR>'' + 
''<TD>QUOTED_IDENTIFIER</TD><TD ALIGN="CENTER">'' + CASE is_quoted_identifier_on WHEN 0 THEN ''<FONT COLOR="BLUE">OFF</FONT>'' ELSE ''<FONT COLOR="GREEN">ON</FONT>'' END + ''</TD>'' + 
''<TD>READ_COMMITTED_SNAPSHOT</TD><TD>'' + CASE is_read_committed_snapshot_on WHEN 0 THEN ''<FONT COLOR="BLUE">OFF</FONT>'' ELSE ''<FONT COLOR="GREEN">ON</FONT>'' END + ''</TD>'' + 
''</TR>'' + 

''<TR>'' + 
''<TD>ANSI_PADDING</TD><TD ALIGN="CENTER">'' + CASE is_ansi_padding_on WHEN 0 THEN ''<FONT COLOR="BLUE">OFF</FONT>'' ELSE ''<FONT COLOR="GREEN">ON</FONT>'' END + ''</TD>'' + 
''<TD>ALLOW_SNAPSHOT_ISOLATION</TD><TD>'' + CASE snapshot_isolation_state WHEN 0 THEN ''<FONT COLOR="BLUE">OFF</FONT>'' ELSE ''<FONT COLOR="GREEN">ON</FONT>'' END + ''</TD>'' + 
''</TR>'' + 

''<TR>'' + 
''<TD>DB Is cleanly shutdown</TD><TD ALIGN="CENTER">'' + CASE is_cleanly_shutdown WHEN 0 THEN ''<FONT COLOR="BLUE">OFF</FONT>'' ELSE ''<FONT COLOR="GREEN">ON</FONT>'' END + ''</TD>'' + 
''<TD>FORCED_PARAMETRIZATION</TD><TD>'' + CASE is_parameterization_forced WHEN 0 THEN ''<FONT COLOR="BLUE">OFF</FONT>'' ELSE ''<FONT COLOR="GREEN">ON</FONT>'' END + ''</TD>'' + 
''</TR>'' + 

''<TR>'' + 
''<TD>FULL-TEXT</TD><TD ALIGN="CENTER">'' + CASE is_fulltext_enabled WHEN 0 THEN ''<FONT COLOR="BLUE">OFF</FONT>'' ELSE ''<FONT COLOR="GREEN">ON</FONT>'' END + ''</TD>'' + 
''<TD>CONCAT_NULL_YIELDS_NULL</TD><TD>'' + CASE is_concat_null_yields_null_on WHEN 0 THEN ''<FONT COLOR="BLUE">OFF</FONT>'' ELSE ''<FONT COLOR="GREEN">ON</FONT>'' END + ''</TD>'' + 
''</TR>'' + 
''<TR>'' + 
''<TD>SUPPLEMENTAL_LOGGING</TD><TD ALIGN="CENTER">'' + CASE is_supplemental_logging_enabled WHEN 0 THEN ''<FONT COLOR="BLUE">OFF</FONT>'' ELSE ''<FONT COLOR="GREEN">ON</FONT>'' END + ''</TD>'' + 
''<TD>Log Reuse</TD><TD>'' + log_reuse_wait_desc + ''</TD>'' + 
''</TR>'' + 
''<TR>'' + 
''<TD>RECURSIVE_TRIGGERS</TD><TD ALIGN="CENTER">'' + CASE is_recursive_triggers_on WHEN 0 THEN ''<FONT COLOR="BLUE">OFF</FONT>'' ELSE ''<FONT COLOR="GREEN">ON</FONT>'' END + ''</TD>'' + 
''<TD>DB_ENCRYPTED_MASTER_KEY</TD><TD>'' + CASE is_master_key_encrypted_by_server WHEN 0 THEN ''<FONT COLOR="BLUE">OFF</FONT>'' ELSE ''<FONT COLOR="GREEN">ON</FONT>'' END + ''</TD>'' + 
''</TR>'' + 
''<TR>'' + 
''<TD>CURSOR_DEFAULT</TD><TD ALIGN="CENTER">'' + CASE is_local_cursor_default WHEN 0 THEN ''<FONT COLOR="BLUE">OFF</FONT>'' ELSE ''<FONT COLOR="GREEN">ON</FONT>'' END + ''</TD>'' + 
''<TD>CURSOR_CLOSE_ON_COMMIT</TD><TD>'' + CASE is_cursor_close_on_commit_on WHEN 0 THEN ''<FONT COLOR="BLUE">OFF</FONT>'' ELSE ''<FONT COLOR="GREEN">ON</FONT>'' END + ''</TD>'' + 
''</TR>'' + 
''<TR>'' + 
''<TD>DB_BROKER</TD><TD ALIGN="CENTER">'' + CASE is_broker_enabled WHEN 0 THEN ''<FONT COLOR="BLUE">OFF</FONT>'' ELSE ''<FONT COLOR="GREEN">ON</FONT>'' END + ''</TD>'' + 
''<TD>HONOR_BROKER_PRIORITY</TD><TD>'' + CASE ' + CASE WHEN @ServerRelease >= 100 THEN 'is_honor_broker_priority_on' ELSE '0' END +
			' WHEN 0 THEN ''<FONT COLOR="BLUE">OFF</FONT>'' ELSE ''<FONT COLOR="GREEN">ON</FONT>'' END + ''</TD>'' + 
''</TR>'' + 
''<TR>'' + 
''<TD>DB_CHAINING</TD><TD ALIGN="CENTER">'' + CASE is_db_chaining_on WHEN 0 THEN ''<FONT COLOR="BLUE">OFF</FONT>'' ELSE ''<FONT COLOR="GREEN">ON</FONT>'' END + ''</TD>'' + 
''<TD>REPLICATION_SUBSCRIPTION_DB</TD><TD>'' + CASE is_subscribed WHEN 0 THEN ''<FONT COLOR="BLUE">OFF</FONT>'' ELSE ''<FONT COLOR="GREEN">ON</FONT>'' END + ''</TD>'' + 
''</TR>'' + 
''<TR>'' + 
''<TD>ARITHABORT</TD><TD ALIGN="CENTER">'' + CASE is_arithabort_on WHEN 0 THEN ''<FONT COLOR="BLUE">OFF</FONT>'' ELSE ''<FONT COLOR="GREEN">ON</FONT>'' END + ''</TD>'' + 
''<TD>REPLICATION_PUBLICATION_DB</TD><TD>'' + CASE is_published WHEN 0 THEN ''<FONT COLOR="BLUE">OFF</FONT>'' ELSE ''<FONT COLOR="GREEN">ON</FONT>'' END + ''</TD>'' + 
''</TR>'' + 
''<TR>'' + 
''<TD>NUMERIC_ROUNDABORT</TD><TD ALIGN="CENTER">'' + CASE is_numeric_roundabort_on WHEN 0 THEN ''<FONT COLOR="BLUE">OFF</FONT>'' ELSE ''<FONT COLOR="GREEN">ON</FONT>'' END + ''</TD>'' + 
''<TD>MERGE_REPLICATION_PUBLICATION_DB</TD><TD>'' + CASE is_merge_published WHEN 0 THEN ''<FONT COLOR="BLUE">OFF</FONT>'' ELSE ''<FONT COLOR="GREEN">ON</FONT>'' END + ''</TD>'' + 
''</TR>'' + 
''<TR>'' + 
''<TD>TRUSTWORTHY</TD><TD ALIGN="CENTER">'' + CASE is_trustworthy_on WHEN 0 THEN ''<FONT COLOR="BLUE">OFF</FONT>'' ELSE ''<FONT COLOR="GREEN">ON</FONT>'' END + ''</TD>'' + 
''<TD>REPLICATION_DISTRIBUTION_DB</TD><TD>'' + CASE is_distributor WHEN 0 THEN ''<FONT COLOR="BLUE">OFF</FONT>'' ELSE ''<FONT COLOR="GREEN">ON</FONT>'' END + ''</TD>'' + 
''</TR>'' + 
''<TR>'' + 
''<TD>AUTO_CLOSE</TD><TD ALIGN="CENTER">'' + CASE is_auto_close_on WHEN 0 THEN ''<FONT COLOR="BLUE">OFF</FONT>'' ELSE ''<FONT COLOR="GREEN">ON</FONT>'' END + ''</TD>'' + 
''<TD>MARKED_FOR_REPLICATION_BACKUP_SYNC</TD><TD>'' + CASE is_sync_with_backup WHEN 0 THEN ''<FONT COLOR="BLUE">OFF</FONT>'' ELSE ''<FONT COLOR="GREEN">ON</FONT>'' END + ''</TD>'' + 
''</TR>'' + 

''<TR>'' + 
''<TD>CHANGE_DATA_CAPTURE</TD><TD ALIGN="CENTER">'' + CASE ' + CASE WHEN @ServerRelease >= 100 THEN 'is_cdc_enabled' ELSE '0' END +
			' WHEN 0 THEN ''<FONT COLOR="BLUE">OFF</FONT>'' ELSE ''<FONT COLOR="GREEN">ON</FONT>'' END + ''</TD>'' + 
''<TD>PAGE_VERIFY</TD><TD style="color:'' + CASE page_verify_option_desc WHEN ''CHECKSUM'' THEN ''GREEN'' ELSE ''RED'' END + '';">'' + page_verify_option_desc + ''</TD>'' + 
''</TR>'' + 

''<TR>'' + 
''<TD>COMPATIBILITY LEVEL</TD><TD ALIGN="CENTER" style="color:'' + 
	CASE WHEN compatibility_level < ' + CAST(@ServerRelease as VARCHAR) + ' THEN ''RED'' ELSE ''GREEN'' END + '';">'' + 
	CASE compatibility_level WHEN 80 THEN ''2000'' WHEN 90 THEN ''2005'' WHEN 100 THEN ''2008'' WHEN 110 THEN ''2012'' WHEN 120 THEN ''2014'' ELSE ''N/A'' END + 
''</TD>'' + 
''<TD>DATE_CORRELATION_OPTIMIZATION</TD><TD>'' + CASE is_date_correlation_on WHEN 0 THEN ''<FONT COLOR="BLUE">OFF</FONT>'' ELSE ''<FONT COLOR="GREEN">ON</FONT>'' END + ''</TD>'' + 
''</TR>'' 

FROM sys.databases
WHERE name = ''' + @DBName + ''';';


  IF @Debug = 1 PRINT @TempString;

  BEGIN TRY
    INSERT INTO @Results(ResultRecord) EXECUTE (@TempString);
  END TRY

  BEGIN CATCH
    SET @ErrorMessage = ERROR_MESSAGE();
    
    BEGIN TRY
        RAISERROR (@ErrorMessage, 16, 0);
    END TRY
    
    BEGIN CATCH
      PRINT @TempString; PRINT ERROR_MESSAGE();
      SELECT @ErrorList = @ErrorList + 
        'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
        'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
    END CATCH
  END CATCH

INSERT INTO @Results(ResultRecord) VALUES ('</TABLE>');

RAISERROR ('#0022 Finished', 0, 1) WITH NOWAIT;


/*0024 Collecting Information About Backups */
SET @CurrentStep = '0025';
INSERT INTO @Results(ResultRecord)
VALUES ('<H3>Database Backups:</H3>' +
  '<Table border=1 cellpadding=5><TR><TH>Backup Type</TH><TH>Last time Taken</TH></TR>' ) ;

SET @TempString = '
	;WITH t as (
		SELECT ''D'' as BU_Type, ''Database'' as BU_Type_Description
		UNION ALL SELECT ''I'', ''Differential database''
		UNION ALL SELECT ''L'', ''Log''
		UNION ALL SELECT ''F'', ''File or filegroup''
		UNION ALL SELECT ''G'', ''Differential file''
		UNION ALL SELECT ''P'', ''Partial''
		UNION ALL SELECT ''Q'', ''Differential partial''
	)
	SELECT ''<TR><TD>'' + t.BU_Type_Description + ''</TD><TD>'' + 
		IsNull(CONVERT(VARCHAR,MAX(backup_finish_date),120),''Never'') + ''</TD></TR>''
	FROM t LEFT JOIN msdb.dbo.backupset as b 
		ON t.BU_Type = b.type and b.Database_name = ''' + @DBName + '''
	GROUP BY t.BU_Type_Description
	ORDER BY t.BU_Type_Description;';

  IF @Debug = 1 PRINT @TempString;

  BEGIN TRY
    INSERT INTO @Results(ResultRecord) EXECUTE (@TempString);
  END TRY

  BEGIN CATCH
    SET @ErrorMessage = ERROR_MESSAGE();
    
    BEGIN TRY
        RAISERROR (@ErrorMessage, 16, 0);
    END TRY
    
    BEGIN CATCH
      PRINT @TempString; PRINT ERROR_MESSAGE();
      SELECT @ErrorList = @ErrorList + 
        'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
        'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
    END CATCH
  END CATCH


INSERT INTO @Results(ResultRecord) VALUES ('</TABLE>');

RAISERROR ('#0021 Finished', 0, 1) WITH NOWAIT;



/*0025 Collecting General Information About File Groups */
SET @CurrentStep = '0025';
INSERT INTO @Results(ResultRecord)
VALUES ('<H3>Database Filegroups:</H3>' +
  '<Table border=1 cellpadding=5><TR><TH>ID</TH><TH>Filegroup Name</TH><TH>FileGroup Type</TH><TH>Is Default</TH><TH>Used for<BR>full-text index</TH><TH>Is Read-Only</TH><TH>Files in<BR>Filegroup</TH><TH>Description</TH></TR>' ) ;

SET @TempString = '
    ;WITH CNT as (SELECT Groupid, COUNT(*) AS "Count" FROM ' + @DBName + '.sys.sysfiles GROUP BY Groupid)
    SELECT ''<TR><TD align="CENTER">'' + CAST(fg.data_space_id as varchar) + ''</TD><TD>'' + fg.name COLLATE ' + @Collation + ' + ''</TD><TD>'' +
      CASE [type] 
        WHEN ''FG'' THEN ''Filegroup''
        WHEN ''PS'' THEN ''Partition scheme''
        WHEN ''FD'' THEN ''FILESTREAM data filegroup''
      END + ''</TD><TD align="CENTER">'' +
      CASE is_default WHEN 1 THEN ''YES'' ELSE ''NO'' END + ''</TD><TD align="CENTER">'' + '
      + CASE WHEN @ServerVersion >= '2012' THEN 'CASE is_system WHEN 1 THEN ''YES'' ELSE ''NO'' END + '
        ELSE '''N/A'' + ' END + 
      '''</TD><TD align="CENTER">'' + 
      CASE is_read_only WHEN 1 THEN ''YES'' ELSE ''NO'' END + ''</TD><TD align="CENTER">'' +
      CAST(IsNull(CNT.[Count],0) as varchar) + ''</TD><TD>'' + IsNull(e.value,''&nbsp'') + ''</TD></TR>''
    FROM ' + @DBName + '.sys.filegroups as fg
    LEFT JOIN CNT ON CNT.Groupid = fg.data_space_id
    LEFT JOIN ##Temp_extended_Properties_' + @SessionId + ' as e ON e.class = 20 and e.major_id = fg.data_space_id;';
  IF @Debug = 1 PRINT @TempString;

  BEGIN TRY
    INSERT INTO @Results(ResultRecord) EXECUTE (@TempString);
  END TRY

  BEGIN CATCH
    SET @ErrorMessage = ERROR_MESSAGE();
    
    BEGIN TRY
        RAISERROR (@ErrorMessage, 16, 0);
    END TRY
    
    BEGIN CATCH
      PRINT @TempString; PRINT ERROR_MESSAGE();
      SELECT @ErrorList = @ErrorList + 
        'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
        'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
    END CATCH
  END CATCH


INSERT INTO @Results(ResultRecord)
VALUES ('</TABLE>');

RAISERROR ('#0025 Finished', 0, 1) WITH NOWAIT;

/*0030 Collecting General Information About Data Files */
SET @CurrentStep = '0030';
INSERT INTO @Results(ResultRecord) VALUES 
	('<H3>Database Files:</H3>');
INSERT INTO @Results(ResultRecord) VALUES 
	('<Table border=1 cellpadding=5><TR><TH>File<BR>ID</TH><TH>File Group</TH><TH>File Name</TH><TH>Physical Name</TH>');
INSERT INTO @Results(ResultRecord) VALUES 
	('<TH>Max File<BR>Size in GB</TH><TH>File<BR>Grow by</TH><TH>File Size<BR>in MB</TH><TH>Used Space<BR>in MB</TH><TH>Description</TH></TR> ' ) ;

SET @TempString = 
'USE [' + @DBName + ']; SELECT ''<TR><TD ALIGN="CENTER">'' + CAST(file_id as VARCHAR) + ''</TD><TD>'' + 
  g.name  COLLATE ' + @Collation + ' + ''</TD><TD>'' + 
  f.name COLLATE ' + @Collation + ' + ''</TD><TD>'' + f.physical_name + ''</TD><TD align="RIGHT">'' + 
  CASE WHEN f.max_size < 0 THEN ''Unlimited'' ELSE  CAST(CAST(ROUND(f.max_size/ (1024. * 128.),3) AS DECIMAL(11,3)) AS VARCHAR) END 
   + ''</TD><TD align="RIGHT">'' + CASE is_percent_growth WHEN 0 THEN CASE 
    WHEN f.growth < 128 THEN CAST(f.growth * 8 AS VARCHAR)  + ''KB'' 
    WHEN f.growth < 131072 THEN CAST(CAST(ROUND(f.growth /128.,3) AS DECIMAL(12,3)) AS VARCHAR) + ''MB''
    ELSE CAST(CAST(ROUND(f.growth /131072.,3) AS DECIMAL(12,3)) AS VARCHAR) + ''GB'' END
    ELSE CAST(f.growth AS VARCHAR)  + ''%''  END + ''</TD><TD align="RIGHT">'' + 
   CAST(CAST(ROUND(f.size/128.,3) AS DECIMAL(11,3)) AS VARCHAR) + ''</TD><TD align="RIGHT">'' + 
  CAST(CAST(ROUND((FILEPROPERTY(f.name, ''SpaceUsed''))/128.,3)  AS DECIMAL(11,3)) AS VARCHAR)  + ''</TD><TD>'' + IsNull(e.value,''&nbsp'') + ''</TD></TR>'' 
FROM sys.database_files as f  INNER JOIN sys.filegroups as g ON f.data_space_id = g.data_space_id
LEFT JOIN ##Temp_extended_Properties_' + @SessionId + ' as e ON e.class = 22 and e.major_id = f.file_id
WHERE f.type_desc != ''LOG''; ';

IF @Debug = 1 PRINT @TempString;

  BEGIN TRY
    INSERT INTO @Results(ResultRecord) EXECUTE (@TempString);
  END TRY

  BEGIN CATCH
    SET @ErrorMessage = ERROR_MESSAGE();
    
    BEGIN TRY
        RAISERROR (@ErrorMessage, 16, 0);
    END TRY
    
    BEGIN CATCH
      PRINT @TempString; PRINT ERROR_MESSAGE();
      SELECT @ErrorList = @ErrorList + 
        'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
        'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
    END CATCH
  END CATCH

RAISERROR ('#0030 Finished', 0, 1) WITH NOWAIT;

/*0040 Collecting General Information About Log Files */
SET @CurrentStep = '0040';

  SET @TempString = 'DBCC LogInfo(''' + @DBName + ''');'
  IF @Debug = 1 PRINT @TempString;

  BEGIN TRY
	IF @ServerVersion >= 2012
		INSERT INTO @LogInfo(RecoveryUnitId, FileId, FileSize, StartOffset, FSeqNo, Status, Parity, CreateLSN)
		EXEC (@TempString);
	ELSE
		INSERT INTO @LogInfo(FileId, FileSize, StartOffset, FSeqNo, Status, Parity, CreateLSN)
		EXEC (@TempString);
	SET @i = @@IDENTITY;
  END TRY

  BEGIN CATCH
    SET @ErrorMessage = ERROR_MESSAGE();
    
    BEGIN TRY
        RAISERROR (@ErrorMessage, 16, 0);
    END TRY
    
    BEGIN CATCH
      PRINT @TempString; PRINT ERROR_MESSAGE();
      SELECT @ErrorList = @ErrorList + 
        'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
        'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
        SET @i = 0;
    END CATCH
  END CATCH

RAISERROR ('#0040 Finished', 0, 1) WITH NOWAIT;


SET @CurrentStep = '0041';
SET @TempString = 'SELECT @IntermediateResults = CAST(CAST(ROUND(cntr_value/1024.,3) AS DECIMAL(11,3)) AS VARCHAR)  FROM sys.dm_os_performance_counters
    WHERE counter_name = ''Log File(s) Used Size (KB)'' AND instance_name = @DBName;'

BEGIN TRY
	/* If user do not have permissions to read Log File size there will be an error resulting "N/A" for Log file size */
	EXEC sp_executesql @TempString, 
		N'@DBName VARCHAR(128), @IntermediateResults VARCHAR(30) OUTPUT', 
		@DBName = @DBName, 
		@IntermediateResults = @IntermediateResults OUTPUT;
END TRY
BEGIN CATCH
	SET @IntermediateResults = 'N/A';
END CATCH

SET @TempString = 
'SELECT ''<TR><TD ALIGN="CENTER">'' + CAST(f.file_id as VARCHAR) + ''</TD><TD>&nbsp</TD><TD>'' + f.name + ''</TD><TD>'' + f.physical_name + ''</TD><TD align="RIGHT">'' + 
  CASE WHEN max_size < 0 THEN ''Unlimited'' ELSE  CAST(CAST(ROUND(max_size/ (1024. * 128.),3) AS DECIMAL(11,3)) AS VARCHAR) END 
   + ''</TD><TD align="RIGHT">'' + CASE is_percent_growth WHEN 0 THEN CASE 
    WHEN growth < 128 THEN CAST(growth * 8 AS VARCHAR)  + ''KB'' 
    WHEN growth < 131072 THEN CAST(CAST(ROUND(growth /128.,3) AS DECIMAL(12,3)) AS VARCHAR) + ''MB''
    ELSE CAST(CAST(ROUND(growth /131072.,3) AS DECIMAL(12,3)) AS VARCHAR) + ''GB'' END
    ELSE CAST(growth AS VARCHAR)  + ''%''  END + ''</TD><TD align="RIGHT">'' + 
   CAST(CAST(ROUND(f.size/128.,3) AS DECIMAL(11,3)) AS VARCHAR) + ''</TD><TD align="RIGHT">'' + ''' +
   @IntermediateResults + ''' + ''</TD><TD ALIGN="CENTER">'' + 
    CASE WHEN ' + CAST(@i as VARCHAR) + ' = 0 THEN ''Error'' ELSE ''# of VLFs: ' + CAST(@i as VARCHAR) + ''' END + ''</TD></TR></TABLE>'' 
  FROM ' + @DBName + '.sys.database_files f  WHERE TYPE = 1; ';

  IF @Debug = 1 PRINT @TempString;
  BEGIN TRY
    INSERT INTO @Results(ResultRecord) EXECUTE (@TempString);
  END TRY

  BEGIN CATCH
    SET @ErrorMessage = ERROR_MESSAGE();
    
    BEGIN TRY
        RAISERROR (@ErrorMessage, 16, 0);
    END TRY
    
    BEGIN CATCH
      PRINT @TempString; PRINT ERROR_MESSAGE();
      SELECT @ErrorList = @ErrorList + 
        'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
        'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
    END CATCH
  END CATCH

RAISERROR ('#0041 Finished', 0, 1) WITH NOWAIT;

/*0050 Collect list of all Objects */
SET @CurrentStep = '0050';
SET @TempString = '
	SELECT sname, schema_id, oname, object_id, parent_object_id, create_date, modify_date, Type, type_desc, EPValue
	FROM (
		SELECT s.name as sname, o.schema_id, o.name as oname, o.object_id, o.parent_object_id, o.create_date, o.modify_date, o.Type, o.type_desc, IsNull(e.value,''&nbsp'') as EPValue,
		CASE Type WHEN ''U'' THEN 0 WHEN ''V'' THEN 15 WHEN ''P'' THEN 10 WHEN ''FN'' THEN 20 WHEN ''TF'' THEN 30 WHEN ''TR'' THEN 40 WHEN ''SN'' THEN 45
		WHEN ''D'' THEN 50 WHEN ''C'' THEN 60 WHEN ''UQ'' THEN 70 WHEN ''PK'' THEN 80 WHEN ''F'' THEN 90 ELSE 200 END as SortOrder
		FROM ' + @DBName + N'.sys.objects as o
			INNER JOIN ' + @DBName + N'.sys.schemas as s ON o.schema_id = s.schema_id
			LEFT JOIN ##Temp_extended_Properties_' + @SessionId + ' as e ON o.object_id = e.major_id and e.minor_id = 0
	  UNION 
	  SELECT ''N/A'', 0, t.name, t.object_id, Null, t.create_date, t.modify_date, t.Type, t.type_desc, IsNull(e.value,''&nbsp''),
		CASE Type WHEN ''U'' THEN 0 WHEN ''V'' THEN 15 WHEN ''P'' THEN 10 WHEN ''FN'' THEN 20 WHEN ''TF'' THEN 30 WHEN ''TR'' THEN 40 WHEN ''SN'' THEN 45
		WHEN ''D'' THEN 50 WHEN ''C'' THEN 60 WHEN ''UQ'' THEN 70 WHEN ''PK'' THEN 80 WHEN ''F'' THEN 90 ELSE 200 END as SortOrder
    FROM ' + @DBName + N'.sys.triggers as t
    LEFT JOIN ##Temp_extended_Properties_' + @SessionId + ' as e ON t.object_id = e.major_id and e.minor_id = 0 and e.class = 1
    WHERE t.parent_class = 0
    ) a ORDER BY SortOrder, oname, sname;';
    
-- ORDER BY CASE ObjectType WHEN 'P' THEN 1 WHEN 'FN' THEN 2 WHEN 'TF' THEN 3 WHEN 'TR' THEN 4 END    
IF @Debug = 1 PRINT @TempString;

  BEGIN TRY
    INSERT INTO @Objects(SchemaName, SchemaID, ObjectName,  ObjectID, ParentObjectID, Created_Dt, Modified_Dt, ObjectType, ObjectType_Desc, Ext_Property) 
    EXECUTE (@TempString);
  END TRY

  BEGIN CATCH
    SET @ErrorMessage = ERROR_MESSAGE();
    
    BEGIN TRY
        RAISERROR (@ErrorMessage, 16, 0);
    END TRY
    
    BEGIN CATCH
      PRINT @TempString; PRINT ERROR_MESSAGE();
      SELECT @ErrorList = @ErrorList + 
        'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
        'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
    END CATCH
  END CATCH

RAISERROR ('#0050 Finished', 0, 1) WITH NOWAIT;

/*0060 Collect list of User Defined Data Types */
SET @CurrentStep = '0060';
SET @TempString = '
    SELECT s.name, u.schema_id, u.name, 0, u.user_type_id, ''UD'', ''USER DEFINED DATA TYPES'', IsNull(CAST(e.value as varchar(256)),''&nbsp'') 
    FROM ' + @DBName + '.sys.types as u 
    INNER JOIN ' + @DBName + '.sys.types as t ON t.user_type_id = u.system_type_id and u.is_user_defined = 1
    INNER JOIN ' + @DBName + '.sys.schemas as s ON u.schema_id = s.schema_id
    LEFT JOIN ##Temp_extended_Properties_' + @SessionId + ' as e ON u.user_type_id = e.major_id and e.minor_id = 0 and e.class = 6
    ORDER BY u.name;';
IF @Debug = 1 PRINT @TempString;


  BEGIN TRY
    INSERT INTO @Objects(SchemaName, SchemaID, ObjectName,  ObjectID, ParentObjectID,ObjectType, ObjectType_Desc, Ext_Property) EXECUTE (@TempString);
  END TRY

  BEGIN CATCH
    SET @ErrorMessage = ERROR_MESSAGE();
    
    BEGIN TRY
        RAISERROR (@ErrorMessage, 16, 0);
    END TRY
    
    BEGIN CATCH
      PRINT @TempString; PRINT ERROR_MESSAGE();
      SELECT @ErrorList = @ErrorList + 
        'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
        'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
    END CATCH
  END CATCH


IF @Debug = 1 SELECT * FROM @Objects ORDER BY ObjectType, ObjectName, ObjectID

RAISERROR ('#0060 Finished', 0, 1) WITH NOWAIT;

/*0065 Collect schema names */
SET @CurrentStep = '0065';
SET @TempString = 
'SELECT ''SCHEMA'', ''SC'', Count(*) FROM ' + @DBName + '.sys.schemas WHERE schema_id < 16384;';
IF @Debug = 1 PRINT @TempString;

  BEGIN TRY
    INSERT INTO @ObjStats(ObjectTypeDesc, ObjectType, ObjCnt) EXECUTE (@TempString);
  END TRY

  BEGIN CATCH
    SET @ErrorMessage = ERROR_MESSAGE();
    
    BEGIN TRY
        RAISERROR (@ErrorMessage, 16, 0);
    END TRY
    
    BEGIN CATCH
      PRINT @TempString; PRINT ERROR_MESSAGE();
      SELECT @ErrorList = @ErrorList + 
        'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
        'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
    END CATCH
  END CATCH

RAISERROR ('#0065 Finished', 0, 1) WITH NOWAIT;

/*0070 Collect  Objects' Statistics */
SET @CurrentStep = '0070';
  BEGIN TRY
    INSERT INTO @ObjStats(ObjectTypeDesc, ObjectType, ObjCnt) 
    SELECT ObjectType_Desc, ObjectType, COUNT(*)
    FROM @Objects 
    GROUP BY  ObjectType, ObjectType_Desc
    ORDER BY (
	    CASE ObjectType
		    WHEN 'U' THEN 0
		    WHEN 'V' THEN 15
		    WHEN 'P' THEN 10
		    WHEN 'FN' THEN 20
		    WHEN 'TF' THEN 30
		    WHEN 'TR' THEN 40
		    WHEN 'SN' THEN 45
		    WHEN 'D' THEN 50
		    WHEN 'C' THEN 60
		    WHEN 'UQ' THEN 70
		    WHEN 'PK' THEN 80
		    WHEN 'F' THEN 90
		    ELSE 200
	    END
    );
  END TRY

  BEGIN CATCH
    SET @ErrorMessage = ERROR_MESSAGE();
    
    BEGIN TRY
        RAISERROR (@ErrorMessage, 16, 0);
    END TRY
    
    BEGIN CATCH
      PRINT @TempString; PRINT ERROR_MESSAGE();
      SELECT @ErrorList = @ErrorList + 
        'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
        'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
    END CATCH
  END CATCH

RAISERROR ('#0070 Finished', 0, 1) WITH NOWAIT;

/*0075 Collect  Number of Partition Objects */
SET @CurrentStep = '0075';
SET @TempString = '
    SELECT ''PARTITION FUNCTIONS'', CASE NAME WHEN '''' THEN '''' ELSE ''PF'' END, COUNT(*) FROM ' + @DBName + '.sys.partition_functions GROUP BY CASE NAME WHEN '''' THEN '''' ELSE ''PF'' END
    UNION ALL
    SELECT ''PARTITION SCHEMES'',  CASE NAME WHEN '''' THEN '''' ELSE ''PS'' END, COUNT(*) FROM ' + @DBName + '.sys.partition_schemes GROUP BY CASE NAME WHEN '''' THEN '''' ELSE ''PS'' END;' ;			
IF @Debug = 1 PRINT @TempString;

  BEGIN TRY
    INSERT INTO @ObjStats(ObjectTypeDesc, ObjectType, ObjCnt) EXECUTE (@TempString);
  END TRY

  BEGIN CATCH
    SET @ErrorMessage = ERROR_MESSAGE();
    
    BEGIN TRY
        RAISERROR (@ErrorMessage, 16, 0);
    END TRY
    
    BEGIN CATCH
      PRINT @TempString; PRINT ERROR_MESSAGE();
      SELECT @ErrorList = @ErrorList + 
        'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#0210)<BR>' +
        'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
    END CATCH
  END CATCH

RAISERROR ('#0075 Finished', 0, 1) WITH NOWAIT;

/*0080 Collect Indexes if any */
SET @CurrentStep = '0080';
SET @TempString = 
'SELECT i.type_desc + '' INDEXES'', ''IX'', Count(*)
FROM ' + @DBName + '.sys.indexes i
INNER JOIN ' + @DBName + '.sys.objects o ON o.object_id = i.object_id
WHERE o.type = ''U'' and i.type_desc IN (''CLUSTERED'', ''NONCLUSTERED'')
GROUP BY i.type_desc;';
IF @Debug = 1 PRINT @TempString;

  BEGIN TRY
    INSERT INTO @ObjStats(ObjectTypeDesc, ObjectType, ObjCnt) EXECUTE (@TempString);
  END TRY

  BEGIN CATCH
    SET @ErrorMessage = ERROR_MESSAGE();
    
    BEGIN TRY
        RAISERROR (@ErrorMessage, 16, 0);
    END TRY
    
    BEGIN CATCH
      PRINT @TempString; PRINT ERROR_MESSAGE();
      SELECT @ErrorList = @ErrorList + 
        'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
        'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
    END CATCH
  END CATCH

RAISERROR ('#0080 Finished', 0, 1) WITH NOWAIT;

/*0090 Check if any XML Schema Collection exist */
SET @CurrentStep = '0090';
SET @TempString = '
    SELECT ''XML SCHEMA COLLECTIONS'', ''XC'', COUNT(*) 
    FROM ' + @DBName + '.sys.xml_schema_collections as x
    INNER JOIN ' + @DBName + '.sys.schemas as s on s.schema_id = x.schema_id
    WHERE x.name != ''sys'' HAVING COUNT(*) > 0;'
IF @Debug = 1 PRINT @TempString;

  BEGIN TRY
    INSERT INTO @ObjStats(ObjectTypeDesc, ObjectType, ObjCnt) EXECUTE (@TempString);
  END TRY

  BEGIN CATCH
    SET @ErrorMessage = ERROR_MESSAGE();
    
    BEGIN TRY
        RAISERROR (@ErrorMessage, 16, 0);
    END TRY
    
    BEGIN CATCH
      PRINT @TempString; PRINT ERROR_MESSAGE();
      SELECT @ErrorList = @ErrorList + 
        'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
        'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
    END CATCH
  END CATCH

INSERT INTO @Results(ResultRecord)
VALUES ('<H2>Objects'' Statistics:</H2>' +
  '<TABLE border=1 cellpadding=5>' +
  '<TR><TH>Object Type</TH><TH>Count</TH>' +
  '<TH>Object Type</TH><TH>Count</TH>' +
  '<TH>Object Type</TH><TH>Count</TH></TR>');

RAISERROR ('#0090 Finished', 0, 1) WITH NOWAIT;

/*0100 Generate content of Object statistic table  */
SET @CurrentStep = '0100';

  BEGIN TRY
    ;WITH Devided AS (
      SELECT *, (ROW_NUMBER() over(order by StatID)+2) % 3  as ObjOrder 
      FROM @ObjStats
    )
    INSERT INTO @Results(ResultRecord) 
    SELECT 
      CASE ObjOrder WHEN 0 THEN '<TR>' ELSE '' END + '<TD>' + 
      CASE WHEN ObjectType IN ('U','P','FN','TF','TR','V','SN','UD','XC','SC','PS','PF') 
      THEN '<A HREF="#' + CASE 
        WHEN ObjectType = 'U' THEN 'Table' 
        WHEN ObjectType = 'P' THEN 'Proc' 
        WHEN ObjectType = 'V' THEN 'View' 
        WHEN ObjectType = 'TR' THEN 'Trig' 
        WHEN ObjectType = 'SN' THEN 'Syn' 
        WHEN ObjectType = 'UD' THEN 'UDDT' 
        WHEN ObjectType = 'XC' THEN 'XSC' 
        WHEN ObjectType = 'SC' THEN 'SCHEMA' 
        WHEN ObjectType = 'PF' THEN 'PartFunc' 
        WHEN ObjectType = 'PS' THEN 'PartSchema' 
        ELSE 'Func' 
      END + 'List">' + ObjectTypeDesc + '</A>'
      ELSE ObjectTypeDesc END + '</TD><TD ALIGN="CENTER">' +
      CAST(ObjCnt AS VARCHAR) + '</TD>' +
      CASE ObjOrder WHEN 2 THEN '</TR>' ELSE '' END
    FROM Devided
    ORDER BY StatID;
  END TRY

  BEGIN CATCH
    SET @ErrorMessage = ERROR_MESSAGE();
    
    BEGIN TRY
        RAISERROR (@ErrorMessage, 16, 0);
    END TRY
    
    BEGIN CATCH
      PRINT @TempString; PRINT ERROR_MESSAGE();
      SELECT @ErrorList = @ErrorList + 
        'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
        'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
    END CATCH
  END CATCH

IF @Debug = 1 SELECT * FROM @ObjStats ORDER BY StatID;

INSERT INTO @Results(ResultRecord) VALUES ('</TABLE>');

RAISERROR ('#0100 Finished', 0, 1) WITH NOWAIT;

/****************************************************************************************************************************************************************************************************************************************/
/* This Section Checks for possible DB design violations */
/* Here are inserted numbers of possible violations. Currently inserted six, but actually used only two*/
INSERT INTO @vt(ViolationID) VALUES (1)
INSERT INTO @vt(ViolationID) VALUES (2)
INSERT INTO @vt(ViolationID) VALUES (3)
INSERT INTO @vt(ViolationID) VALUES (4)
INSERT INTO @vt(ViolationID) VALUES (5)
INSERT INTO @vt(ViolationID) VALUES (6)

/*0110 Check for tables without clustered indexes - Violation #1 */
  SET @CurrentStep = '0110';
  SET @TempString = '
    SELECT @i = COUNT(*) FROM ' + @DBName + '.sys.indexes as i
    INNER JOIN ' + @DBName + '.sys.tables as t ON i.object_id = t.object_id
    WHERE INDEX_ID = 0;';
  IF @Debug = 1 PRINT @TempString;
  
  BEGIN TRY
    EXEC sp_executesql @TempString, N'@i int OUTPUT', @i = @i OUTPUT 
  END TRY

  BEGIN CATCH
    SET @ErrorMessage = ERROR_MESSAGE();
    
    BEGIN TRY
        RAISERROR (@ErrorMessage, 16, 0);
    END TRY
    
    BEGIN CATCH
      PRINT @TempString; PRINT ERROR_MESSAGE();
      SELECT @ErrorList = @ErrorList + 
        'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
        'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
    END CATCH
  END CATCH


  IF @Debug = 1 SELECT @i as "Tables without Clustered Index"
  /* Update violation table for Violation #1 */
  UPDATE @vt SET Violated = @i  WHERE ViolationID = 1

RAISERROR ('#0110 Finished', 0, 1) WITH NOWAIT;

/*0120 Check for possible duplicate indexes - Violation #2 */
  SET @CurrentStep = '0120';
  SET @TempString = '
      ;WITH ForResearch AS (
          SELECT  o.object_id, s.name as schemaname, o.name as tablename, i.name as IndexName, 
              ic.column_id, ic.key_ordinal, i.index_id, i.type, IsNull(x.secondary_type_desc, ''PRIMARY'') as XML_Type
          FROM ' + @DBName + '.sys.indexes i
          INNER JOIN ' + @DBName + '.sys.objects o ON i.object_id = o.object_id
          INNER JOIN ' + @DBName + '.sys.schemas s ON s.schema_id = o.schema_id
          INNER JOIN ' + @DBName + '.sys.index_columns ic ON ic.index_id = i.index_id and ic.object_id = o.object_id
          INNER JOIN ' + @DBName + '.sys.columns c ON ic.column_id = c.column_id and c.object_id = o.object_id
          LEFT JOIN ' + @DBName + '.sys.xml_indexes as x ON x.index_id = i.index_id and o.object_id = x.object_id
          WHERE i.index_id > 0 and s.name != ''sys''
      )
      SELECT @i = COUNT(DISTINCT t3.schemaname + t3.tablename + t3.IndexName)
      FROM ForResearch  as t3
      WHERE Not exists (
      SELECT t1.object_id FROM ForResearch as t1
      LEFT JOIN ForResearch as t2 on t1.object_id = t2.object_id  and t1.column_id = t2.column_id  and 
        t1.index_id != t2.index_id and t1.XML_Type = t2.XML_Type and 
        ( (t1.key_ordinal != 1 and t2.key_ordinal != 1) or t1.key_ordinal = t2.key_ordinal )
      WHERE t2.object_id Is Null and t1.object_id = t3.object_id and t1.index_id = t3.index_id);';

  
  IF @Debug = 1 PRINT @TempString;
  BEGIN TRY
    EXEC sp_executesql @TempString, N'@i int OUTPUT', @i = @i OUTPUT 
  END TRY

  BEGIN CATCH
    SET @ErrorMessage = ERROR_MESSAGE();
    
    BEGIN TRY
        RAISERROR (@ErrorMessage, 16, 0);
    END TRY
    
    BEGIN CATCH
      PRINT @TempString; PRINT ERROR_MESSAGE();
      SELECT @ErrorList = @ErrorList + 
        'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
        'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
    END CATCH
  END CATCH

  IF @Debug = 1 SELECT @i as "Tables without Clustered Index"
  /* Update violation table for Violation #2 */
  UPDATE @vt SET Violated = @i  WHERE ViolationID = 2

/*0130 Report violations if any */
SET @CurrentStep = '0130';
IF EXISTS (SELECT * FROM  @vt WHERE Violated > 0)
BEGIN
  INSERT INTO @Results(ResultRecord) VALUES ('<H2>Alert Box:</H2>' +
    '<TABLE border=1 cellpadding=5>' +
    '<TR><TH>Possible Violation</TH><TH>Number</TH></TR>');

  /*0140 Listing  tables without clustered indexes if any */
    IF EXISTS (SELECT * FROM  @vt WHERE Violated > 0 and ViolationID = 1)
    BEGIN
      INSERT INTO @Results(ResultRecord) 
      SELECT TOP 1 '<TR><TD><A HREF="#NoClusteredIndex">Tables Without Clustered Index</A></TD><TD align="CENTER">' + CAST(Violated AS VARCHAR)  + '</TD></TR>'
      FROM @vt WHERE ViolationID = 1;
    END 

  /*0150 Listing possible duplicate indexes if any */
    IF EXISTS (SELECT * FROM  @vt WHERE Violated > 0 and ViolationID = 2)
    BEGIN
      INSERT INTO @Results(ResultRecord) 
      SELECT TOP 1 '<TR><TD><A HREF="#DuplicateIndex">Index Duplication</A></TD><TD align="CENTER">' + CAST(Violated AS VARCHAR)  + '</TD></TR>'
      FROM @vt WHERE ViolationID = 2;
    END 

  INSERT INTO @Results(ResultRecord) VALUES ('</TABLE>');
END

RAISERROR ('#0120 Finished', 0, 1) WITH NOWAIT;

/****************************************************************************************************************************************************************************************************************************************/
/* Listing Partitioning objects */

/*0210 Listing all partitioning Functions, if any */
SET @CurrentStep = '0210';

IF EXISTS ( SELECT TOP 1 1 FROM @ObjStats WHERE ObjectType = 'PF')
BEGIN
  INSERT INTO @Results(ResultRecord) VALUES 
    ('<H2><A id="PartFuncList">List of Partition Functions:</A></H2>' +
    '<TABLE border=1 cellpadding=5><TR><TH>Partition<BR>Schema</TH>' +
    '<TH>Partition<BR>Function</TH><TH>Partition<BR>Function<BR>Type</TH>' +
    '<TH>Boundary<BR>Direction</TH><TH>Boundary<BR>Type</TH>' +
    '<TH>Boundary<BR>ID</TH><TH>Partition<BR>Function<BR>Value</TH>' +
    '<TH>File<BR>Group</TH><TH>Function Description</TH>' +
    '<TH>Function<BR>Creation<BR>Date</TH><TH>Function<BR>Modification<BR>Date</TH></TR>');
  
  SET @TempString = '
    SELECT ''<TR><TD ALIGN="CENTER">'' + ps.name + ''</TD><TD>'' +  pf.name + ''</TD><TD>'' + 
      pf.type_desc COLLATE ' + @Collation + ' + ''</TD><TD>'' +
      CASE pf.boundary_value_on_right WHEN 0 THEN ''LEFT'' ELSE ''RIGHT'' END + 
      ''</TD><TD>'' + st.name + 
      ''</TD><TD ALIGN="CENTER">'' + CAST(prv.boundary_id as VARCHAR) + 
      ''</TD><TD ALIGN="CENTER">'' + fg.name + 
      ''</TD><TD ALIGN="CENTER">'' + CAST(prv.value as VARCHAR) + ''</TD><TD>'' +
      IsNull(e.value,''&nbsp'') + ''</TD><TD>'' + 
      CAST(pf.Create_Date as VARCHAR) + ''</TD><TD>'' + 
      CAST(pf.Modify_Date as VARCHAR) + ''</TD></TR>''
    FROM ' + @DBName + '.sys.partition_functions AS pf 
    INNER JOIN ' + @DBName + '.sys.partition_parameters as pp ON pf.function_id = pp.function_id
    INNER JOIN ' + @DBName + '.sys.types as st ON st.system_type_id = pp.system_type_id
    INNER JOIN ' + @DBName + '.sys.partition_schemes as ps ON pf.function_id  = ps.function_id 
    INNER JOIN ' + @DBName + '.sys.destination_data_spaces as ds ON ds.partition_scheme_id = ps.data_space_id
    INNER JOIN ' + @DBName + '.sys.filegroups as fg ON ds.data_space_id = fg.data_space_id
    INNER JOIN ' + @DBName + '.sys.partition_range_values as prv ON prv.function_id=pf.function_id and prv.boundary_id = ds.destination_id
    LEFT JOIN ##Temp_extended_Properties_' + @SessionId + ' as e ON e.class = 21 and e.major_id = pf.function_id
    ORDER BY pf.name, prv.boundary_id;';

  IF @Debug = 1 PRINT @TempString;

  BEGIN TRY
    INSERT INTO @Results(ResultRecord) EXECUTE (@TempString);  
  END TRY

  BEGIN CATCH
    SET @ErrorMessage = ERROR_MESSAGE();
    
    BEGIN TRY
        RAISERROR (@ErrorMessage, 16, 0);
    END TRY
    
    BEGIN CATCH
      PRINT @TempString; PRINT ERROR_MESSAGE();
      SELECT @ErrorList = @ErrorList + 
        'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
        'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
    END CATCH
  END CATCH

  INSERT INTO @Results(ResultRecord) VALUES ('</TABLE>');
END

RAISERROR ('#0210 Finished', 0, 1) WITH NOWAIT;

/*0220 Listing all partitioning Schemas, if any */
SET @CurrentStep = '0220';
IF EXISTS ( SELECT TOP 1 1 FROM @ObjStats WHERE ObjectType = 'PS')
BEGIN
  INSERT INTO @Results(ResultRecord) VALUES 
    ('<H2><A id="PartSchemaList">List of Partition Schemes:</A></H2>' +
    '<TABLE border=1 cellpadding=5>' +
    '<TR><TH>Partition<BR>Scheme</TH><TH>Partition<BR>Function</TH>' +
    '<TH>Partition<BR>Function<BR>Type</TH><TH>Scheme Description</TH></TR>');
  
  SET @TempString = '
    SELECT ''<TR><TD>'' + ps.name + ''</TD><TD>'' + pf.name + ''</TD><TD>'' + 
      pf.type_desc COLLATE ' + @Collation + ' + ''</TD><TD>'' +
      IsNull(e.value,''&nbsp'') + ''</TD></TR>''
    FROM ' + @DBName + '.sys.partition_functions AS pf 
    INNER JOIN ' + @DBName + '.sys.partition_schemes as ps ON ps.function_id = pf.function_id
    LEFT JOIN ##Temp_extended_Properties_' + @SessionId + ' as e ON e.class = 20 and e.major_id = ps.data_space_id
    ORDER BY ps.name;';

  IF @Debug = 1 PRINT @TempString;

  BEGIN TRY
    INSERT INTO @Results(ResultRecord) EXECUTE (@TempString);  
  END TRY

  BEGIN CATCH
    SET @ErrorMessage = ERROR_MESSAGE();
    
    BEGIN TRY
        RAISERROR (@ErrorMessage, 16, 0);
    END TRY
    
    BEGIN CATCH
      PRINT @TempString; PRINT ERROR_MESSAGE();
      SELECT @ErrorList = @ErrorList + 
        'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
        'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
    END CATCH
  END CATCH

  INSERT INTO @Results(ResultRecord) VALUES ('</TABLE>');
END

RAISERROR ('#0220 Finished', 0, 1) WITH NOWAIT;

/*0230 Listing all partitioned Tables/Views/Indexes, if any */
SET @CurrentStep = '0230';
SET @TempString = '
    SELECT @i = COUNT(*) FROM ' + @DBName + '.sys.indexes i
    INNER JOIN ' + @DBName + '.sys.partition_schemes ps ON ps.data_space_id = i.data_space_id;' ;			
IF @Debug = 1 PRINT @TempString;
EXEC sp_executesql @TempString, N'@i int OUTPUT', @i = @i OUTPUT 

IF @i > 0
BEGIN
  INSERT INTO @Results(ResultRecord) VALUES 
	('<DIV  id="ShowListPartObjects"><H2>List of Partitioned Objects (<a href="javascript:HideShowCode(0,''ShowListPartObjects'');HideShowCode(1,''HideListPartObjects'');"> SHOW </a>)</H2></DIV>');
  INSERT INTO @Results(ResultRecord) VALUES 
	('<DIV style="display:none;" id="HideListPartObjects"><H2>List of Partitioned Objects (<a href="javascript:HideShowCode(0,''HideListPartObjects'');HideShowCode(1,''ShowListPartObjects'');"> HIDE </a>)</H2>');
  INSERT INTO @Results(ResultRecord) VALUES 
    ('<TABLE border=1 cellpadding=5>'+
    '<TR><TH>Partition<BR>Schema</TH><TH>Partition<BR>Function</TH>' +
    '<TH>Object<BR>Schema</TH><TH>Object<BR>Name</TH>' +
    '<TH>Index<BR>Name</TH><TH>Object<BR>Type</TH>' +
    '<TH>Object<BR>Compression</TH><TH>FileGroup<BR>Name</TH>' +
    '<TH>Rows in<BR>Partition</TH><TH>Total<BR>Size,<BR>MB</TH>' +
    '<TH>Used<BR>Size,<BR>MB</TH><TH>Data<BR>Size,<BR>MB</TH></TR>');
  
  SET @TempString = '
    SELECT ''<TR><TD>'' + ps.Name + ''</TD><TD>'' + pf.name + ''</TD><TD>'' + 
      s.name + ''</TD><TD><A HREF="#oid'' + CAST(o.object_id as VARCHAR) + ''">'' + 
      o.name + ''</A></TD><TD>'' + IsNull(i.name, ''HEAP'') + ''</TD><TD>'' + 
      AU.type_desc COLLATE ' + @Collation + ' +
      ''</TD><TD ALIGN="Center">'' + pa.data_compression_Desc + ''</TD><TD>'' + FG.name + 
      ''</TD><TD ALIGN="Right">'' + CAST(pa.rows as VARCHAR) + 
      ''</TD><TD ALIGN="Right">'' + CAST(AU.total_pages / 128 as VARCHAR) + 
      ''</TD><TD ALIGN="Right">'' + CAST(AU.used_pages / 128 as VARCHAR) + 
      ''</TD><TD ALIGN="Right">'' + CAST(AU.data_pages / 128 as VARCHAR) + ''</TD></TR>''
    FROM ' + @DBName + '.sys.indexes i
    INNER JOIN ' + @DBName + '.sys.partition_schemes ps on ps.data_space_id = i.data_space_id
    INNER JOIN ' + @DBName + '.sys.partition_functions pf on pf.function_id = ps.function_id 
    INNER JOIN ' + @DBName + '.sys.partitions AS PA 
      ON PA.object_id = i.object_id AND PA.index_id = i.index_id 
    INNER JOIN ' + @DBName + '.sys.allocation_units AS AU 
      ON (AU.type IN (1, 3) AND AU.container_id = PA.hobt_id) 
        OR (AU.type = 2 AND AU.container_id = PA.partition_id) 
    INNER JOIN ' + @DBName + '.sys.objects AS o ON i.object_id = o.object_id 
    INNER JOIN ' + @DBName + '.sys.schemas AS s ON o.schema_id = s.schema_id 
    INNER JOIN ' + @DBName + '.sys.filegroups AS FG ON FG.data_space_id = AU.data_space_id
    ORDER BY ps.Name, s.Name, o.Name, i.Name, FG.name;';
  
  IF @Debug = 1 PRINT @TempString;
  BEGIN TRY
    INSERT INTO @Results(ResultRecord) EXECUTE (@TempString);
  END TRY

  BEGIN CATCH
    SET @ErrorMessage = ERROR_MESSAGE();
    
    BEGIN TRY
        RAISERROR (@ErrorMessage, 16, 0);
    END TRY
    
    BEGIN CATCH
      PRINT @TempString; PRINT ERROR_MESSAGE();
      SELECT @ErrorList = @ErrorList + 
        'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
        'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
    END CATCH
  END CATCH
  
  INSERT INTO @Results(ResultRecord) VALUES ('</TABLE></DIV>');
END

RAISERROR ('#0230 Finished', 0, 1) WITH NOWAIT;

/****************************************************************************************************************************************************************************************************************************************/

/*1000 Listing Objects */
SET @CurrentStep = '1000';
IF EXISTS (SELECT TOP 1 1 FROM @ObjStats WHERE ObjectType NOT IN ('IT','S','SQ') )
  INSERT INTO @Results(ResultRecord) VALUES ('<H2>List of Objects:</H2>'); 

RAISERROR ('#1000 Finished', 0, 1) WITH NOWAIT;

/*1004 Listing Schemas */
SET @CurrentStep = '1004';

    INSERT INTO @Results(ResultRecord) VALUES 
      ('<H2><A id="SCHEMAList">Database Schemas:</A></H2>'+
      '<TABLE border=1 cellpadding=5>' +
      '<TR><TH>ID</TH><TH>Schema Name</TH><TH>Objects in schema</TH><TH>Description</TH></TR>');

    SET @TempString = '
      ;WITH SchemaInfo as (
        SELECT s.schema_id, s.name COLLATE ' + @Collation + ' as name, Count(o.schema_id) as Cnt,  IsNull(e.value,''&nbsp'') as Value
        FROM ' + @DBName + '.sys.schemas as s 
        LEFT JOIN ' + @DBName + '.sys.objects as o ON s.schema_id = o.schema_id
        LEFT JOIN ##Temp_extended_Properties_' + @SessionId + ' as e ON e.major_id = s.schema_id and e.class = 3
        WHERE s.schema_id < 16384
        GROUP BY s.schema_id, s.Name, e.value
      )  
      SELECT ''<TR><TD ALIGN="CENTER">'' + CAST(schema_id as varchar) + ''</TD><TD>'' + name  + ''</TD><TD ALIGN="CENTER">'' + CAST(Cnt as varchar) + ''</TD><TD>'' + Value  + ''</TD></TR>''
      FROM SchemaInfo ORDER BY schema_id;'

    IF @Debug = 1 PRINT @TempString;
    BEGIN TRY
      INSERT INTO @Results(ResultRecord) EXECUTE (@TempString);
    END TRY

    BEGIN CATCH
      SET @ErrorMessage = ERROR_MESSAGE();
      
      BEGIN TRY
          RAISERROR (@ErrorMessage, 16, 0);
      END TRY
      
      BEGIN CATCH
        PRINT @TempString; PRINT ERROR_MESSAGE();
        SELECT @ErrorList = @ErrorList + 
          'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
          'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
      END CATCH
    END CATCH

    
    INSERT INTO @Results(ResultRecord) VALUES ('</TABLE>');

RAISERROR ('#1004 Finished', 0, 1) WITH NOWAIT;

/*1007 Listing DB logins */
SET @CurrentStep = '1007';

    INSERT INTO @Results(ResultRecord) VALUES 
      ('<A id="TableList"></a>');
    INSERT INTO @Results(ResultRecord) VALUES 
	  ('<DIV  id="ShowDBLoginList"><H2>Database Logins (<a href="javascript:HideShowCode(0,''ShowDBLoginList'');HideShowCode(1,''HideDBLoginList'');"> SHOW </a> )</H2></DIV>');
    INSERT INTO @Results(ResultRecord) VALUES 
	  ('<DIV style="display:none;" id="HideDBLoginList"><H2>Database Logins (<a href="javascript:HideShowCode(0,''HideDBLoginList'');HideShowCode(1,''ShowDBLoginList'');"> HIDE </a>)</H2>');
    INSERT INTO @Results(ResultRecord) VALUES 
      ('<TABLE border=1 cellpadding=5>');
    INSERT INTO @Results(ResultRecord) VALUES 
      ('<TR><TH>Login Name</TH><TH>Login Type</TH><TH>Default<BR>Schema</TH>');
    INSERT INTO @Results(ResultRecord) VALUES 
      ('<TH>Date<BR>Created</TH><TH>Date<BR>Modified</TH><TH>Has DB<BR>Access</TH>');
    INSERT INTO @Results(ResultRecord) VALUES 
      ('<TH>DB<BR>Owner</TH><TH>Access<BR>Admin</TH><TH>Security<BR>Admin</TH>');
    INSERT INTO @Results(ResultRecord) VALUES 
      ('<TH>DDL<BR>Admin</TH><TH>Data<BR>Reader</TH><TH>Data<BR>Writer</TH></TR>');
      
	SET @TempString = '
	;WITH DBUsers as (
		SELECT IsNull(dr.name,'''') as DBRole, dm.name as DBUser, 1 as pvt, 
			dm.principal_id, dm.type_desc, dm.default_schema_name
		FROM [' + @DBName + '].sys.database_principals as dm
		LEFT JOIN [' + @DBName + '].sys.database_role_members as rm on rm.member_principal_id = dm.principal_id
		LEFT JOIN [' + @DBName + '].sys.database_principals as dr on rm.role_principal_id = dr.principal_id
		WHERE dm.is_fixed_role = 0
	),  pvt as (
		SELECT DBUser, principal_id, type_desc, default_schema_name, 
			IsNull([db_owner],0) as [db_owner], 
			IsNull([db_accessadmin],0) as [db_accessadmin], 
			IsNull([db_securityadmin],0) as [db_securityadmin],
			IsNull([db_ddladmin],0) as [db_ddladmin], 
			IsNull([db_datareader],0) as [db_datareader], 
			IsNull([db_datawriter],0) as [db_datawriter]
		FROM DBUsers PIVOT (	
			MAX(pvt) FOR DBRole in (
				[db_owner], [db_accessadmin], [db_securityadmin], 
				[db_ddladmin], [db_datareader], [db_datawriter]
			)
		) As pvt
	)
	SELECT 
		''<TR><TD>'' + pvt.DBUser COLLATE ' + @Collation + ' + 
		''</TD><TD>'' + pvt.type_desc + 
		 ''</TD><TD ALIGN="CENTER">'' + IsNull(pvt.default_schema_name,''&nbsp'') +
		 ''</TD><TD>'' + CONVERT(VARCHAR,s.createdate,120) +
		  ''</TD><TD>'' + CONVERT(VARCHAR,s.updatedate,120) +
		  ''</TD><TH>'' + CASE s.hasdbaccess WHEN 0 THEN ''<FONT COLOR="RED">NO</FONT>'' ELSE ''<FONT COLOR="GREEN">YES</FONT>'' END +  
		  ''</TH><TH>'' + CASE pvt.[db_owner] WHEN 1 THEN ''<FONT COLOR="BLUE">X</FONT>'' ELSE ''&nbsp'' END +
		  ''</TH><TH>'' + CASE pvt.[db_accessadmin] WHEN 1 THEN ''<FONT COLOR="BLUE">X</FONT>'' ELSE ''&nbsp'' END +
		  ''</TH><TH>'' + CASE pvt.[db_securityadmin] WHEN 1 THEN ''<FONT COLOR="BLUE">X</FONT>'' ELSE ''&nbsp'' END +
		  ''</TH><TH>'' + CASE pvt.[db_ddladmin] WHEN 1 THEN ''<FONT COLOR="BLUE">X</FONT>'' ELSE ''&nbsp'' END +
		  ''</TH><TH>'' + CASE pvt.[db_datareader] WHEN 1 THEN ''<FONT COLOR="BLUE">X</FONT>'' ELSE ''&nbsp'' END +
		  ''</TH><TH>'' + CASE pvt.[db_datawriter] WHEN 1 THEN ''<FONT COLOR="BLUE">X</FONT>'' ELSE ''&nbsp'' END + ''</TH></TR>''
	FROM pvt INNER JOIN [' + @DBName + '].sys.sysusers	as s ON s.uid = pvt.principal_id
	ORDER BY DBUser;';

    IF @Debug = 1 PRINT @TempString;
    BEGIN TRY
      INSERT INTO @Results(ResultRecord) EXECUTE (@TempString);
    END TRY

    BEGIN CATCH
      SET @ErrorMessage = ERROR_MESSAGE();
      
      BEGIN TRY
          RAISERROR (@ErrorMessage, 16, 0);
      END TRY
      
      BEGIN CATCH
        PRINT @TempString; PRINT ERROR_MESSAGE();
        SELECT @ErrorList = @ErrorList + 
          'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
          'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
      END CATCH
    END CATCH

    
    INSERT INTO @Results(ResultRecord) VALUES ('</TABLE></DIV>');

RAISERROR ('#1007 Finished', 0, 1) WITH NOWAIT;

/*1010 Listing Tables */
SET @CurrentStep = '1010';
IF EXISTS (SELECT TOP 1 1 FROM @ObjStats WHERE ObjectType = 'U' )
BEGIN
    INSERT INTO @Results(ResultRecord) VALUES 
    ('<A id="TableList"></a>');
    INSERT INTO @Results(ResultRecord) VALUES 
	('<DIV  id="ShowTableList"><H2>User Tables (<a href="javascript:HideShowCode(0,''ShowTableList'');HideShowCode(1,''HideTableList'');"> SHOW </a> )</H2></DIV>');
    INSERT INTO @Results(ResultRecord) VALUES 
	('<DIV style="display:none;" id="HideTableList"><H2>User Tables (<a href="javascript:HideShowCode(0,''HideTableList'');HideShowCode(1,''ShowTableList'');"> HIDE </a>)</H2>');
    INSERT INTO @Results(ResultRecord) VALUES 
      ('<TABLE border=1 cellpadding=5>');
    INSERT INTO @Results(ResultRecord) VALUES 
      ('<TR><TH>##</TH><TH>Schema</TH><TH>Table Name</TH>');
    INSERT INTO @Results(ResultRecord) VALUES 
      ('<TH>Object ID</TH><TH>Created</TH><TH>Modified</TH>');
    INSERT INTO @Results(ResultRecord) VALUES 
      ('<TH>Columns</TH><TH >Row Max<BR>Size (Bytes)</TH><TH>Row Count</TH>');
    INSERT INTO @Results(ResultRecord) VALUES 
      ('<TH>Table Size<BR>in MB</TH><TH>Used Space<BR>in MB</TH>');
    INSERT INTO @Results(ResultRecord) VALUES 
      ('<TH>Compres<BR>sion</TH><TH>Description</TH></TR>');

    SET @TempString = '
    ;WITH TableInfo AS (
      SELECT s.name as SchemaName, t.name as TableName, t.object_id, t.create_date, t.modify_date, t.lob_data_space_id, t.max_column_id_used, 
        SUM(CASE c.max_length WHEN -1 THEN 0 ELSE c.max_length END) as max_length, ROW_NUMBER() over(order by t.name, t.object_id)  as TableNumber
      FROM ' + @DBName + '.sys.tables as t
      INNER JOIN ' + @DBName + '.sys.schemas as s ON t.schema_id = s.schema_id
      INNER JOIN ' + @DBName + '.sys.columns as c ON c.object_id = t.object_id
      /*INNER JOIN ' + @DBName + '.sys.dm_db_partition_stats as st ON t.object_id = st.object_id*/
      GROUP BY s.name, t.name, t.object_id, t.create_date, t.modify_date, t.lob_data_space_id, t.max_column_id_used
    ), Sizes AS (
		SELECT object_id, 
			SUM(CASE WHEN Index_id > 1 THEN 0 ELSE row_count END) as row_count, 
			ROUND(CAST(SUM(reserved_page_count) AS float)/128.,3) as SizeMB, 
			ROUND(CAST(SUM(used_page_count) AS float)/128.,3) AS UsedSpaceMB 
		FROM ' + @DBName + '.sys.dm_db_partition_stats GROUP BY object_id
    )
    SELECT 
      ''<TR><TD>'' + CAST(TableNumber as varchar) + ''</TD><TD ALIGN="CENTER">'' + SchemaName COLLATE ' + @Collation + ' + 
      ''</TD><TD><A HREF="#oid'' + CAST(i.object_id as varchar) + ''" onclick="javascript:HideShowCode(0,''''ShowTablesViews'''');HideShowCode(1,''''HideTablesViews'''');">'' + TableName + ''</A>'' + 
      ''</TD><TD>'' + CAST(i.object_id as varchar) + 
      ''</TD><TD>'' + CAST(create_date as varchar) + ''</TD><TD>'' + CAST(modify_date as varchar) + 
      ''</TD><TD ALIGN="CENTER">'' + CAST(max_column_id_used as varchar) + 
      ''</TD><TD ALIGN="CENTER">'' + CAST(max_length as varchar) + CASE lob_data_space_id WHEN 0 THEN '''' ELSE '' + LOB'' END + 
      ''</TD><TD ALIGN="RIGHT">'' + CAST(s.row_count as varchar) + 
      ''</TD><TD ALIGN="RIGHT">'' + CAST(s.SizeMB as varchar) + 
      ''</TD><TD ALIGN="RIGHT">'' + CAST(s.UsedSpaceMB as varchar) + 
      ''</TD><TD ALIGN="RIGHT">'' + p.data_compression_desc + ''</TD><TD>'' + IsNull(e.value,''&nbsp'') + ''</TD></TR>''
    FROM TableInfo as i INNER JOIN Sizes as s on i.object_id = s.object_id
    LEFT JOIN (SELECT DISTINCT  object_id, ' + 
	CASE WHEN @ServerRelease >= 100 THEN '' ELSE '''N/A'' as ' END +
	' data_compression_desc FROM ' + @DBName + '.sys.partitions WHERE index_id < 2) as p 
		ON p.object_id = i.object_id
    LEFT JOIN ##Temp_extended_Properties_' + @SessionId + ' as e ON e.class = 1 and e.major_id = i.object_id and e.minor_id = 0
    ORDER BY SchemaName, TableName, i.object_id;';

    IF @Debug = 1 PRINT @TempString;
    BEGIN TRY
      INSERT INTO @Results(ResultRecord) EXECUTE (@TempString);
    END TRY

    BEGIN CATCH
      SET @ErrorMessage = ERROR_MESSAGE();
      
      BEGIN TRY
          RAISERROR (@ErrorMessage, 16, 0);
      END TRY
      
      BEGIN CATCH
        PRINT @TempString; PRINT ERROR_MESSAGE();
        SELECT @ErrorList = @ErrorList + 
          'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
          'ErrorMessage: "' + ERROR_MESSAGE() + '" (Most probably no rights to query "sys.dm_db_partition_stats") <BR><BR>';
      END CATCH
    END CATCH
    
    INSERT INTO @Results(ResultRecord) VALUES ('</TABLE></DIV>');
END

RAISERROR ('#1010 Finished', 0, 1) WITH NOWAIT;

/*1020 Listing Views */
SET @CurrentStep = '1020';
IF EXISTS (SELECT TOP 1 1 FROM @ObjStats WHERE ObjectType = 'V' )
BEGIN
    INSERT INTO @Results(ResultRecord) VALUES 
    ('<A id="ViewList"></a>');
    INSERT INTO @Results(ResultRecord) VALUES 
	('<DIV  id="ShowViewList"><H2>User Views (<a href="javascript:HideShowCode(0,''ShowViewList'');HideShowCode(1,''HideViewList'');"> SHOW </a> )</H2></DIV>');
    INSERT INTO @Results(ResultRecord) VALUES 
	('<DIV style="display:none;" id="HideViewList"><H2>User Views (<a href="javascript:HideShowCode(0,''HideViewList'');HideShowCode(1,''ShowViewList'');"> HIDE </a>)</H2>');
    INSERT INTO @Results(ResultRecord) VALUES 
      ('<TABLE border=1  cellpadding=5>');
    INSERT INTO @Results(ResultRecord) VALUES 
      ('<TR><TH>##</TH><TH>Schema</TH><TH>View Name</TH>' + 
      '<TH>Object ID</TH><TH>Created</TH><TH>Modified</TH>' +
      '<TH>Columns</TH><TH>Max Size (Bytes)</TH><TH>Description</TH></TR>');

    SET @TempString = ' 
    ;WITH ViewInfo AS (
      SELECT s.name as SchemaName, v.name as ViewName, v.object_id, v.create_date, v.modify_date,
        SUM(CASE c.max_length WHEN -1 THEN 0 ELSE c.max_length END) as max_length, 
        MIN(c.max_length) as MinLength, COUNT(*) as ColumnNumber,
        ROW_NUMBER() over(order by v.name, v.object_id)  as ViewNumber
      FROM ' + @DBName + '.sys.views as v
      INNER JOIN ' + @DBName + '.sys.schemas as s ON v.schema_id = s.schema_id
      INNER JOIN ' + @DBName + '.sys.columns as c ON c.object_id = v.object_id
      GROUP BY s.name, v.name, v.object_id, v.create_date, v.modify_date
    )
      SELECT 
      ''<TR><TD>'' + CAST(ViewNumber as varchar) + ''</TD><TD ALIGN="CENTER">'' + SchemaName COLLATE ' + @Collation + ' + 
      ''</TD><TD><A HREF="#oid'' + CAST(object_id as varchar) + ''"  onclick="javascript:HideShowCode(0,''''ShowTablesViews'''');HideShowCode(1,''''HideTablesViews'''');">'' + ViewName + 
      ''</A></TD><TD>'' + CAST(object_id as varchar) + 
      ''</TD><TD>'' + CAST(create_date as varchar) + ''</TD><TD>'' + CAST(modify_date as varchar) + ''</TD><TD ALIGN="CENTER">'' + CAST(ColumnNumber as varchar) + 
      ''</TD><TD ALIGN="CENTER">'' + CAST(max_length as varchar) + CASE MinLength WHEN -1 THEN '' + LOB'' ELSE '''' END + ''</TD><TD>'' + IsNull(e.value,''&nbsp'') + ''</TD></TR>''
      FROM ViewInfo as i
      LEFT JOIN ##Temp_extended_Properties_' + @SessionId + ' as e ON e.class = 1 and e.major_id = i.object_id and e.minor_id = 0
      ORDER BY SchemaName, ViewName, object_id;';
    
    IF @Debug = 1 PRINT @TempString;
    BEGIN TRY
      INSERT INTO @Results(ResultRecord) EXECUTE (@TempString);
    END TRY

    BEGIN CATCH
      SET @ErrorMessage = ERROR_MESSAGE();
      
      BEGIN TRY
          RAISERROR (@ErrorMessage, 16, 0);
      END TRY
      
      BEGIN CATCH
        PRINT @TempString; PRINT ERROR_MESSAGE();
        SELECT @ErrorList = @ErrorList + 
          'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
          'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
      END CATCH
    END CATCH
    
    INSERT INTO @Results(ResultRecord) VALUES ('</TABLE></DIV>');
END

RAISERROR ('#1020 Finished', 0, 1) WITH NOWAIT;

/*1030 Listing Procedures */
SET @CurrentStep = '1030';
IF EXISTS (SELECT TOP 1 1 FROM @ObjStats WHERE ObjectType = 'P' )
BEGIN
    INSERT INTO @Results(ResultRecord) VALUES 
    ('<A id="ProcList"></A>');
    INSERT INTO @Results(ResultRecord) VALUES 
	('<DIV  id="ShowProcList"><H2>User Stored Procedures (<a href="javascript:HideShowCode(0,''ShowProcList'');HideShowCode(1,''HideProcList'');"> SHOW </a> )</H2></DIV>');
    INSERT INTO @Results(ResultRecord) VALUES 
	('<DIV style="display:none;" id="HideProcList"><H2>User Stored Procedures (<a href="javascript:HideShowCode(0,''HideProcList'');HideShowCode(1,''ShowProcList'');"> HIDE </a>)</H2>');
    INSERT INTO @Results(ResultRecord) VALUES 
      ('<TABLE border=1  cellpadding=5>');
    INSERT INTO @Results(ResultRecord) VALUES 
      ('<TR><TH>##</TH><TH>Schema</TH><TH>Procedure Name</TH>' + 
      '<TH>Object ID</TH><TH>Created</TH><TH>Modified</TH><TH>Description</TH></TR>');

    SET @TempString = ' 
    ;WITH ProcInfo AS (
      SELECT s.name as SchemaName, p.name as ProcName, p.object_id, p.create_date, p.modify_date,
        ROW_NUMBER() over(order by p.name, p.object_id)  as ProcNumber
      FROM ' + @DBName + '.sys.procedures as p
      INNER JOIN ' + @DBName + '.sys.schemas as s ON p.schema_id = s.schema_id
    )
      SELECT 
      ''<TR><TD ALIGN="CENTER">'' + CAST(ProcNumber as varchar) + ''</TD><TD ALIGN="CENTER">'' + SchemaName COLLATE ' + @Collation + ' + 
      ''</TD><TD><A HREF="#oid'' + CAST(object_id as varchar) + ''"  onclick="javascript:HideShowCode(0,''''ShowProgrammedObjects'''');HideShowCode(1,''''HideProgrammedObjects'''');">'' + ProcName + 
      ''</A></TD><TD>'' + CAST(object_id as varchar) + 
      ''</TD><TD>'' + CAST(create_date as varchar) + ''</TD><TD>'' + CAST(modify_date as varchar) + ''</TD><TD>'' + IsNull(e.value,''&nbsp'') + ''</TD></TR>''
      FROM ProcInfo as i
      LEFT JOIN ##Temp_extended_Properties_' + @SessionId + ' as e ON e.class = 1 and e.major_id = i.object_id and e.minor_id = 0
      ORDER BY SchemaName, ProcName, object_id;';
    
    IF @Debug = 1 PRINT @TempString;
    BEGIN TRY
      INSERT INTO @Results(ResultRecord) EXECUTE (@TempString);
    END TRY

    BEGIN CATCH
      SET @ErrorMessage = ERROR_MESSAGE();
      
      BEGIN TRY
          RAISERROR (@ErrorMessage, 16, 0);
      END TRY
      
      BEGIN CATCH
        PRINT @TempString; PRINT ERROR_MESSAGE();
        SELECT @ErrorList = @ErrorList + 
          'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
          'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
      END CATCH
    END CATCH

    INSERT INTO @Results(ResultRecord) VALUES ('</TABLE></DIV>');
END

RAISERROR ('#1030 Finished', 0, 1) WITH NOWAIT;

/*1040 Listing Functions */
SET @CurrentStep = '1040';
IF EXISTS (SELECT TOP 1 1 FROM @ObjStats WHERE ObjectType IN ('TF','FN') )
BEGIN
    INSERT INTO @Results(ResultRecord) VALUES 
    ('<A id="FuncList"></A>');
    INSERT INTO @Results(ResultRecord) VALUES 
	('<DIV  id="ShowFuncList"><H2>User Functions (<a href="javascript:HideShowCode(0,''ShowFuncList'');HideShowCode(1,''HideFuncList'');"> SHOW </a> )</H2></DIV>');
    INSERT INTO @Results(ResultRecord) VALUES 
	('<DIV style="display:none;" id="HideFuncList"><H2>User Functions (<a href="javascript:HideShowCode(0,''HideFuncList'');HideShowCode(1,''ShowFuncList'');"> HIDE </a>)</H2>');
    INSERT INTO @Results(ResultRecord) VALUES 
      ('<TABLE border=1  cellpadding=5>' +
      '<TR><TH>##</TH><TH>Schema</TH><TH>Function Name</TH><TH>Function Type</TH>' + 
      '<TH>Object ID</TH><TH>Created</TH><TH>Modified</TH><TH>Description</TH></TR>');

    SET @TempString = ' 
    ;WITH FuncInfo AS (
      SELECT s.name as SchemaName, f.name as FuncName, f.object_id, f.create_date, f.modify_date, f.type, 
        CASE f.type_desc WHEN ''SQL_SCALAR_FUNCTION'' THEN ''Scalar'' 
          WHEN ''SQL_TABLE_VALUED_FUNCTION'' THEN ''Table Valued''
          WHEN ''AGGREGATE_FUNCTION'' THEN ''Aggregate''
          WHEN ''CLR_SCALAR_FUNCTION'' THEN ''CLR Scalar''
          WHEN ''CLR_TABLE_VALUED_FUNCTION'' THEN ''CLR Table Valued''
          WHEN ''SQL_INLINE_TABLE_VALUED_FUNCTION'' THEN ''Inline Table Valued''
        ELSE ''Other'' END AS type_desc,
        ROW_NUMBER() over(order by f.type, f.name, f.object_id)  as FuncNumber
      FROM ' + @DBName + '.sys.objects as f
      INNER JOIN ' + @DBName + '.sys.schemas as s ON f.schema_id = s.schema_id
      WHERE f.type in (''FN'',''TF'')
    )
      SELECT 
      ''<TR><TD>'' + CAST(FuncNumber as varchar) + ''</TD><TD ALIGN="CENTER">'' + SchemaName COLLATE ' + @Collation + ' + 
      ''</TD><TD><A HREF="#oid'' + CAST(object_id as varchar) + ''"  onclick="javascript:HideShowCode(0,''''ShowProgrammedObjects'''');HideShowCode(1,''''HideProgrammedObjects'''');">'' + FuncName +
      ''</A></TD><TD>'' + type_desc + ''</TD><TD>'' + CAST(object_id as varchar) + 
      ''</TD><TD>'' + CAST(create_date as varchar) + ''</TD><TD>'' + CAST(modify_date as varchar) + ''</TD><TD>'' + IsNull(e.value,''&nbsp'') + ''</TD></TR>''
      FROM FuncInfo as i
      LEFT JOIN ##Temp_extended_Properties_' + @SessionId + ' as e ON e.class = 1 and e.major_id = i.object_id and e.minor_id = 0
      ORDER BY type, FuncName, object_id;';
    
    IF @Debug = 1 PRINT @TempString;
    BEGIN TRY
      INSERT INTO @Results(ResultRecord) EXECUTE (@TempString);
    END TRY

    BEGIN CATCH
      SET @ErrorMessage = ERROR_MESSAGE();
      
      BEGIN TRY
          RAISERROR (@ErrorMessage, 16, 0);
      END TRY
      
      BEGIN CATCH
        PRINT @TempString; PRINT ERROR_MESSAGE();
        SELECT @ErrorList = @ErrorList + 
          'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
          'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
      END CATCH
    END CATCH

    INSERT INTO @Results(ResultRecord) VALUES ('</TABLE></DIV>');
END

RAISERROR ('#1040 Finished', 0, 1) WITH NOWAIT;

/*1050 Listing Triggers */
SET @CurrentStep = '1050';
IF EXISTS (SELECT TOP 1 1 FROM @ObjStats WHERE ObjectType = 'TR' )
BEGIN
    INSERT INTO @Results(ResultRecord) VALUES 
    ('<A id="TrigList"></A>');
    INSERT INTO @Results(ResultRecord) VALUES 
	('<DIV  id="ShowTrigList"><H2>User Triggers (<a href="javascript:HideShowCode(0,''ShowTrigList'');HideShowCode(1,''HideTrigList'');"> SHOW </a> )</H2></DIV>');
    INSERT INTO @Results(ResultRecord) VALUES 
	('<DIV style="display:none;" id="HideTrigList"><H2>User Triggers (<a href="javascript:HideShowCode(0,''HideTrigList'');HideShowCode(1,''ShowTrigList'');"> HIDE </a>)</H2>');
    INSERT INTO @Results(ResultRecord) VALUES 
      ('<TABLE border=1  cellpadding=5>');
    INSERT INTO @Results(ResultRecord) VALUES 
      ('<TR><TH>##</TH><TH>Schema</TH><TH>Trigger Name</TH><TH>Object ID</TH>' + 
      '<TH>Created</TH><TH>Modified</TH>' + 

      '<TH>Is<BR>Update</TH><TH>Is<BR>Delete</TH>' + 
      '<TH>Is<BR>Insert</TH><TH>Is<BR>After</TH>' + 

      '<TH>Is<BR>Disabled</TH><TH>Instead<BR>Of</TH>' + 
      '<TH>Trigger Type</TH><TH>Parent Object</TH>' + 
      '<TH>Parent Object<BR>Type</TH><TH>Description</TH></TR>');
-- trisupdate, trisdelete, trisinsert, trisafter
    SET @TempString = ' 
		USE [' + @DBName + ']; 
    WITH TrigInfo as (			
	    SELECT IsNull(s.name,''N/A'') as SchemaName, t.name as TrigName, t.object_id, t.parent_class,
	      CASE t.parent_class WHEN 0 THEN ''DATABASE'' ELSE ''On OBJECT'' END as Trigger_Type,
	      ISNULL(o.name, CASE t.parent_class WHEN 0 THEN ''DB'' ELSE ''N/A'' END) as Parent_Object, 
	      t.parent_id, 
        CASE o.type_desc WHEN ''USER_TABLE'' THEN ''Table'' WHEN ''VIEW'' THEN ''View'' ELSE ''N/A'' END as Parent_Type, 
	      t.create_date,  t.modify_date,
	      CASE OBJECTPROPERTY(t.object_id, ''ExecIsUpdateTrigger'') WHEN 1 THEN ''Y'' ELSE ''-'' END as isupdate,
	      CASE OBJECTPROPERTY(t.object_id, ''ExecIsDeleteTrigger'') WHEN 1 THEN ''Y'' ELSE ''-'' END as isdelete,
	      CASE OBJECTPROPERTY(t.object_id, ''ExecIsInsertTrigger'') WHEN 1 THEN ''Y'' ELSE ''-'' END as isinsert,
	      CASE OBJECTPROPERTY(t.object_id, ''ExecIsAfterTrigger'') WHEN 1 THEN ''Y'' ELSE ''-'' END as isafter,
	      CASE t.is_disabled WHEN 0 THEN ''-'' ELSE ''Y'' END as Is_disabled,
	      CASE t.is_instead_of_trigger WHEN 0 THEN ''-'' ELSE ''Y'' END as is_instead_of,
	      ROW_NUMBER() over(order by t.parent_class, t.name, t.object_id)  as TrigNumber
	    FROM ' + @DBName + '.sys.triggers as t 
	    LEFT JOIN ' + @DBName + '.sys.objects as o ON o.object_id = t.parent_id
	    LEFT JOIN ' + @DBName + '.sys.schemas as s ON o.schema_id = s.schema_id
    )			
    SELECT ''<TR><TD>''  + CAST(TrigNumber as varchar) + ''</TD><TD ALIGN="CENTER">'' + SchemaName COLLATE ' + @Collation + ' + 
      ''</TD><TD><A HREF="#oid'' + CAST(object_id as varchar)  + ''"  onclick="javascript:HideShowCode(0,''''ShowProgrammedObjects'''');HideShowCode(1,''''HideProgrammedObjects'''');">'' + TrigName + 
      ''</A></TD><TD>'' + CAST(object_id as varchar) + ''</TD><TD>'' + CAST(create_date as varchar) + 
      ''</TD><TD>'' + CAST(modify_date as varchar) + 
      ''</TD><TD ALIGN="CENTER">'' + isupdate + ''</TD><TD ALIGN="CENTER">'' + isdelete + 
      ''</TD><TD ALIGN="CENTER">'' + isinsert + ''</TD><TD ALIGN="CENTER">'' + isafter + 
      ''</TD><TD ALIGN="CENTER">'' + Is_disabled + ''</TD><TD ALIGN="CENTER">'' + is_instead_of + 
      ''</TD><TD>'' + Trigger_Type COLLATE ' + @Collation + ' + ''</TD><TD>'' + 
      CASE parent_class WHEN 0 THEN Parent_Object ELSE ''<A HREF="#oid'' + CAST(parent_id as varchar)  + ''">'' + Parent_Object + ''</A>'' END +
       ''</TD><TD>'' + Parent_Type + ''</TD><TD>'' + IsNull(e.value,''&nbsp'') + ''</TD></TR>''
    FROM TrigInfo as i
    LEFT JOIN ##Temp_extended_Properties_' + @SessionId + ' as e ON e.class = 1 and e.major_id = i.object_id and e.minor_id = 0;';			
    
    IF @Debug = 1 PRINT @TempString;
    BEGIN TRY
      INSERT INTO @Results(ResultRecord) EXECUTE (@TempString);
    END TRY

    BEGIN CATCH
      SET @ErrorMessage = ERROR_MESSAGE();
      
      BEGIN TRY
          RAISERROR (@ErrorMessage, 16, 0);
      END TRY
      
      BEGIN CATCH
        PRINT @TempString; PRINT ERROR_MESSAGE();
        SELECT @ErrorList = @ErrorList + 
          'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
          'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
      END CATCH
    END CATCH

    INSERT INTO @Results(ResultRecord) VALUES ('</TABLE></DIV>');
END 

RAISERROR ('#1050 Finished', 0, 1) WITH NOWAIT;

/*1060 Listing Synonyms */
SET @CurrentStep = '1060';
IF EXISTS (SELECT TOP 1 1 FROM @ObjStats WHERE ObjectType = 'SN' )
BEGIN
    INSERT INTO @Results(ResultRecord) VALUES 
		('<A id="SynList"></A>');
    INSERT INTO @Results(ResultRecord) VALUES 
		('<DIV  id="ShowSynList"><H2>Database Synonyms (<a href="javascript:HideShowCode(0,''ShowSynList'');HideShowCode(1,''HideSynList'');"> SHOW </a> )</H2></DIV>');
    INSERT INTO @Results(ResultRecord) VALUES 
		('<DIV style="display:none;" id="HideSynList"><H2>Database Synonyms (<a href="javascript:HideShowCode(0,''HideSynList'');HideShowCode(1,''ShowSynList'');"> HIDE </a>)</H2>');
    INSERT INTO @Results(ResultRecord) VALUES 
      ('<TABLE border=1  cellpadding=5>');
    INSERT INTO @Results(ResultRecord) VALUES 
      ('<TR><TH>##</TH><TH>Schema Name</TH><TH>Synonym Name</TH>' + 
      '<TH>Created</TH><TH>Modified</TH>' + 
      '<TH>Base Object Name</TH></TR>');

    SET @TempString = '
        SELECT ''<TR><TD ALIGN="CENTER">'' + CAST(ROW_NUMBER() over(order by y.name, s.name) as VARCHAR) + ''</TD><TD>'' + 
          s.name + ''</TD><TD>'' + y.name + ''</TD><TD>'' + CAST(y.create_date as VARCHAR) + ''</TD><TD>'' + 
          CAST(y.modify_date as VARCHAR) + ''</TD><TD>'' + y.base_object_name + ''</TD></TR>''
        FROM ' + @DBName + '.sys.synonyms as y
        INNER JOIN ' + @DBName + '.sys.schemas as s on s.schema_id = y.schema_id
        ORDER BY y.name, s.name;';
    IF @Debug = 1 PRINT @TempString;
    INSERT INTO @Results(ResultRecord) EXECUTE (@TempString);
    INSERT INTO @Results(ResultRecord) VALUES ('</TABLE></DIV>');
END 

RAISERROR ('#1060 Finished', 0, 1) WITH NOWAIT;

/*1070 Listing User Defined Data Types */
SET @CurrentStep = '1070';
IF EXISTS (SELECT TOP 1 1 FROM @ObjStats WHERE ObjectType = 'UD' )
BEGIN
    INSERT INTO @Results(ResultRecord) VALUES 
		('<A id="UDDTList"></A>');
    INSERT INTO @Results(ResultRecord) VALUES 
		('<DIV  id="ShowUDDTList"><H2>User Defined Data Types (<a href="javascript:HideShowCode(0,''ShowUDDTList'');HideShowCode(1,''HideUDDTList'');"> SHOW </a> )</H2></DIV>');
    INSERT INTO @Results(ResultRecord) VALUES 
		('<DIV style="display:none;" id="HideUDDTList"><H2>User Defined Data Types (<a href="javascript:HideShowCode(0,''HideUDDTList'');HideShowCode(1,''ShowUDDTList'');"> HIDE </a>)</H2>');
    INSERT INTO @Results(ResultRecord) VALUES 
      ('<H2><A id="UDDTList">User Defined Data Types:</A></H2>');
    INSERT INTO @Results(ResultRecord) VALUES 
      ('<TABLE border=1  cellpadding=5>');
    INSERT INTO @Results(ResultRecord) VALUES 
      ('<TR><TH>##</TH><TH>User Type</TH><TH>System Type</TH>' + 
      '<TH>Max Length</TH><TH>Precision</TH><TH>Scale</TH>' + 
      '<TH>Collation</TH></TR>');

    SET @TempString = '
        SELECT ''<TR><TD ALIGN="CENTER">'' + CAST(ROW_NUMBER() over(order by t.name, u.name) as VARCHAR) + ''</TD><TD>'' + 
          ''<A HREF="#UDDT'' + CAST(u.user_type_id AS VARCHAR) + ''">'' +  u.name + ''</A></TD><TD>'' + t.name + ''</TD><TD ALIGN="CENTER">'' + 
          CAST(u.max_length/(CASE WHEN t.name IN (''NVARCHAR'',''NCHAR'') THEN 2 ELSE 1 END) as VARCHAR) + 
          ''</TD><TD ALIGN="CENTER">'' + CAST(u.precision as VARCHAR) + ''</TD><TD ALIGN="CENTER">'' + 
          CAST(u.scale as VARCHAR) + ''</TD><TD>'' + IsNull(u.collation_name,''N/A'') + ''</TD></TR>''
         FROM ' + @DBName + '.sys.types as u 
         INNER JOIN ' + @DBName + '.sys.types as t ON t.user_type_id = u.system_type_id and u.is_user_defined = 1
         ORDER BY t.name, u.name;'
    
    IF @Debug = 1 PRINT @TempString;
    BEGIN TRY
      INSERT INTO @Results(ResultRecord) EXECUTE (@TempString);
    END TRY

    BEGIN CATCH
      SET @ErrorMessage = ERROR_MESSAGE();
      
      BEGIN TRY
          RAISERROR (@ErrorMessage, 16, 0);
      END TRY
      
      BEGIN CATCH
        PRINT @TempString; PRINT ERROR_MESSAGE();
        SELECT @ErrorList = @ErrorList + 
          'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
          'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
      END CATCH
    END CATCH

    INSERT INTO @Results(ResultRecord) VALUES ('</TABLE></DIV>');
END

RAISERROR ('#1070 Finished', 0, 1) WITH NOWAIT;

/*1080 Check if any XML Schema Collection exist */
SET @CurrentStep = '1080';
IF EXISTS (SELECT TOP 1 1 FROM @ObjStats WHERE ObjectType = 'XC' )
BEGIN
    INSERT INTO @Results(ResultRecord) VALUES 
		('<A id="XSCList"></A>');
    INSERT INTO @Results(ResultRecord) VALUES 
		('<DIV  id="ShowXSCList"><H2>XML Schema Collections (<a href="javascript:HideShowCode(0,''ShowXSCList'');HideShowCode(1,''HideXSCList'');"> SHOW </a> )</H2></DIV>');
    INSERT INTO @Results(ResultRecord) VALUES 
		('<DIV style="display:none;" id="HideXSCList"><H2>XML Schema Collections (<a href="javascript:HideShowCode(0,''HideXSCList'');HideShowCode(1,''ShowXSCList'');"> HIDE </a>)</H2>');
    INSERT INTO @Results(ResultRecord) VALUES 
      ('<TABLE border=1  cellpadding=5>');
    INSERT INTO @Results(ResultRecord) VALUES 
      ('<TR><TH>##</TH><TH>Schema Name</TH><TH>XML Schema Collection</TH>' + 
      '<TH>Created</TH><TH>Modified</TH></TR>');

    SET @TempString = '
        SELECT ''<TR><TD ALIGN="CENTER">'' + CAST(ROW_NUMBER() over(order by s.name, x.name) as VARCHAR) + ''</TD><TD>'' + 
          s.name + ''</TD><TD>'' + x.name + ''</TD><TD ALIGN="CENTER">'' + CAST(x.create_date as VARCHAR) + 
          ''</TD><TD ALIGN="CENTER">'' + CAST(x.modify_date as VARCHAR) + ''</TD></TR>''
         FROM ' + @DBName + '.sys.xml_schema_collections as x
         INNER JOIN ' + @DBName + '.sys.schemas as s on s.schema_id = x.schema_id
         WHERE x.name != ''sys''
         ORDER BY s.name, x.name;'

    IF @Debug = 1 PRINT @TempString;
    BEGIN TRY
      INSERT INTO @Results(ResultRecord) EXECUTE (@TempString);
    END TRY

    BEGIN CATCH
      SET @ErrorMessage = ERROR_MESSAGE();
      
      BEGIN TRY
          RAISERROR (@ErrorMessage, 16, 0);
      END TRY
      
      BEGIN CATCH
        PRINT @TempString; PRINT ERROR_MESSAGE();
        SELECT @ErrorList = @ErrorList + 
          'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
          'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
      END CATCH
    END CATCH

    INSERT INTO @Results(ResultRecord) VALUES ('</TABLE></DIV>');
END

RAISERROR ('#1080 Finished', 0, 1) WITH NOWAIT;

/********************************************************************************************************************************************************************************************************************/
/* Listings of violation if any */
/*1090 Tables without Clustered Index */
  SET @CurrentStep = '1090';
  IF EXISTS (SELECT * FROM  @vt WHERE Violated > 0 and ViolationID = 1)
  BEGIN
    
    INSERT INTO @Results(ResultRecord) VALUES 
		('<A id="NoClusteredIndex"></A>');
    INSERT INTO @Results(ResultRecord) VALUES 
		('<DIV  id="ShowNCIList"><H2>Tables without Clustered Index (<a href="javascript:HideShowCode(0,''ShowNCIList'');HideShowCode(1,''HideNCIList'');"> SHOW </a> )</H2></DIV>');
    INSERT INTO @Results(ResultRecord) VALUES 
		('<DIV style="display:none;" id="HideNCIList"><H2>Tables without Clustered Index (<a href="javascript:HideShowCode(0,''HideNCIList'');HideShowCode(1,''ShowNCIList'');"> HIDE </a>)</H2>');
    INSERT INTO @Results(ResultRecord) VALUES 
      ('<TABLE border=1 cellpadding=5>' + 
      '<TR><TH>##</TH><TH>Schema</TH>' + 
      '<TH>Table Name</TH><TH>Object Id</TH></TR>'
      );
    SET @TempString = ' 
      ;WITH TableInfo as (
        SELECT DISTINCT t.name as TableName, s.name as SchemaName, i.object_id,
        ROW_NUMBER() over(order by t.name, s.name)  as TableNumber
        FROM ' + @DBName + '.sys.indexes as i
        INNER JOIN ' + @DBName + '.sys.tables as t ON i.object_id = t.object_id 
        INNER JOIN ' + @DBName + '.sys.schemas as s ON s.schema_id = t.schema_id
        WHERE i.INDEX_ID = 0 
      )
      SELECT ''<TR><TD>'' + CAST(TableNumber as varchar)  + ''</TD><TD ALIGN="CENTER">'' + [SchemaName] + 
        ''</TD><TD><A HREF="#oid'' + CAST(object_id as varchar) + ''"  onclick="javascript:HideShowCode(0,''''ShowTablesViews'''');HideShowCode(1,''''HideTablesViews'''');">'' + [TableName]  +
        ''</TD><TD>'' + CAST(object_id as varchar) + ''</TD></TR>''
      FROM TableInfo
      ORDER BY [TableName], [SchemaName];';			

    IF @Debug = 1 PRINT @TempString;
    BEGIN TRY
      INSERT INTO @Results(ResultRecord) EXECUTE (@TempString);
    END TRY

    BEGIN CATCH
      SET @ErrorMessage = ERROR_MESSAGE();
      
      BEGIN TRY
          RAISERROR (@ErrorMessage, 16, 0);
      END TRY
      
      BEGIN CATCH
        PRINT @TempString; PRINT ERROR_MESSAGE();
        SELECT @ErrorList = @ErrorList + 
          'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
          'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
      END CATCH
    END CATCH

    INSERT INTO @Results(ResultRecord) VALUES ('</TABLE></DIV>');
  END

RAISERROR ('#1090 Finished', 0, 1) WITH NOWAIT;

/*1100 Listing possible duplicate indexes if any */
  SET @CurrentStep = '1100';
  IF EXISTS (SELECT * FROM  @vt WHERE Violated > 0 and ViolationID = 2)
  BEGIN
    
    INSERT INTO @Results(ResultRecord) VALUES
		('<A id="DuplicateIndex"></A>');
    INSERT INTO @Results(ResultRecord) VALUES
		('<DIV  id="ShowDblIxList"><H2>Possible Duplicate Indexes (<a href="javascript:HideShowCode(0,''ShowDblIxList'');HideShowCode(1,''HideDblIxList'');"> SHOW </a> )</H2></DIV>');
    INSERT INTO @Results(ResultRecord) VALUES
		('<DIV style="display:none;" id="HideDblIxList"><H2>Possible Duplicate Indexes (<a href="javascript:HideShowCode(0,''HideDblIxList'');HideShowCode(1,''ShowDblIxList'');"> HIDE </a>)</H2>');
    INSERT INTO @Results(ResultRecord) VALUES
      ('<TABLE border=1 cellpadding=5>' + 
      '<TR><TH>##</TH><TH>Schema</TH>' + 
      '<TH>Table Name</TH><TH>Index Name</TH><TH>Object Id</TH></TR>'
      );
    SET @TempString = '
        ;WITH ForResearch AS (
            SELECT  o.object_id, s.name as schemaname, o.name as tablename, i.name as IndexName, 
                ic.column_id, ic.key_ordinal, i.index_id, i.type, IsNull(x.secondary_type_desc, ''PRIMARY'') as XML_Type
            FROM ' + @DBName + '.sys.indexes i
            INNER JOIN ' + @DBName + '.sys.objects o ON i.object_id = o.object_id
            INNER JOIN ' + @DBName + '.sys.schemas s ON s.schema_id = o.schema_id
            INNER JOIN ' + @DBName + '.sys.index_columns ic ON ic.index_id = i.index_id and ic.object_id = o.object_id
            INNER JOIN ' + @DBName + '.sys.columns c ON ic.column_id = c.column_id and c.object_id = o.object_id
            LEFT JOIN ' + @DBName + '.sys.xml_indexes as x ON x.index_id = i.index_id and o.object_id = x.object_id
            WHERE i.index_id > 0 and s.name != ''sys''
        ),
        Results AS (
            SELECT DISTINCT t3.schemaname, t3.tablename, t3.IndexName, t3.object_id
            FROM ForResearch  as t3
            WHERE Not exists (
            SELECT t1.object_id FROM ForResearch as t1
            LEFT JOIN ForResearch as t2 on t1.object_id = t2.object_id  and t1.column_id = t2.column_id  and 
              t1.index_id != t2.index_id and t1.XML_Type = t2.XML_Type and 
              ( (t1.key_ordinal != 1 and t2.key_ordinal != 1) or t1.key_ordinal = t2.key_ordinal )
            WHERE t2.object_id Is Null and t1.object_id = t3.object_id and t1.index_id = t3.index_id)
        )
        SELECT 
            ''<TR><TD>'' + CAST(ROW_NUMBER() over(order by tablename, schemaname)  AS VARCHAR )+ ''</TD><TD ALIGN="CENTER">'' + 
            schemaname + ''</TD><TD><A HREF="#oid'' + CAST(object_id as varchar) + ''"  onclick="javascript:HideShowCode(0,''''ShowTablesViews'''');HideShowCode(1,''''HideTablesViews'''');">'' + 
            tablename + ''</TD><TD>'' + IndexName + ''</TD><TD>'' + CAST(object_id AS VARCHAR)  + ''</TD></TR>''
        FROM Results ORDER BY tablename, schemaname;';

    IF @Debug = 1 PRINT @TempString;
    BEGIN TRY
      INSERT INTO @Results(ResultRecord) EXECUTE (@TempString);
    END TRY

    BEGIN CATCH
      SET @ErrorMessage = ERROR_MESSAGE();
      
      BEGIN TRY
          RAISERROR (@ErrorMessage, 16, 0);
      END TRY
      
      BEGIN CATCH
        PRINT @TempString; PRINT ERROR_MESSAGE();
        SELECT @ErrorList = @ErrorList + 
          'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
          'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
      END CATCH
    END CATCH

    INSERT INTO @Results(ResultRecord) VALUES ('</TABLE></DIV>');
  END  

RAISERROR ('#1100 Finished', 0, 1) WITH NOWAIT;

/********************************************************************************************************************************************************************************************************************/

IF @ReportOnlyObjectNames = 0
BEGIN
    INSERT INTO @Results(ResultRecord) VALUES
	('<DIV  id="ShowTablesViews"><H2>Tables'' and Views'' Details (<a href="javascript:HideShowCode(0,''ShowTablesViews'');HideShowCode(1,''HideTablesViews'');"> SHOW </a>)</H2></DIV>');
    INSERT INTO @Results(ResultRecord) VALUES
	('<DIV style="display:none;" id="HideTablesViews"><H2>Tables'' and Views'' Details (<a href="javascript:HideShowCode(0,''HideTablesViews'');HideShowCode(1,''ShowTablesViews'');"> HIDE </a>)</H2>')

SELECT @TotalNumber = COuNT(*), @CurrentNumber = 0, @NextMilestone = 0 
FROM @Objects 
WHERE ObjectType IN ('U','V', 'FN','P','TR','TF');

IF @TotalNumber > 0
BEGIN
	SET @CntMessage = 'Discovered ' + CAST(@TotalNumber as varchar) + ' of Database objects. Start looping through them.' 
	RAISERROR (@CntMessage, 0, 1) WITH NOWAIT;
	SET @IncrementCount = @TotalNumber / CASE
		WHEN @TotalNumber > @5Percent THEN 100
		WHEN @TotalNumber < @10Percent THEN 10
	ELSE 20 END;
	SET @NextMilestone = @IncrementCount;
END

IF EXISTS (SELECT TOP 1 1 FROM @Objects WHERE ObjectType in ('U','V') and Reported = 0)
BEGIN
/* Generate list of all columns in Database */
	SET @TempString = '
			SELECT c.object_id, c.column_id, c.name, c.user_type_id, UPPER(usrt.name) as Type_Name,
				CASE c.is_identity WHEN 0 THEN ''N'' ELSE ''Y'' END as is_identity,
				CASE usrt.is_user_defined WHEN 0 THEN ''N'' ELSE ''Y'' END as is_user_defined,
				CAST(CAST(CASE WHEN baset.name IN (N''nchar'', N''nvarchar'') AND c.max_length <> -1 THEN c.max_length/2 ELSE c.max_length END AS int) as varchar) as max_length,
				CASE c.is_nullable WHEN 0 THEN ''N'' ELSE ''Y'' END as is_nullable,
				IsNull(c.Collation_Name COLLATE ' + @Collation + ',''N/A'') as Collation_Name,
				CASE c.is_computed WHEN 1 THEN ''Y'' ELSE ''N'' END as is_computed,
				IsNull(cc.definition,''N/A'') as definition,
				CASE cc.is_persisted WHEN 1 THEN ''Y'' WHEN 0 THEN ''N'' ELSE ''N/A'' END as is_persisted,
				IsNull(e.value,''&nbsp'') as ExtValue
			INTO ##Object_Columns_' + @SessionId + '
			FROM ' + @DBName + '.sys.objects as o
			INNER JOIN ' + @DBName + '.sys.columns as c ON c.object_id = o.object_id 
				LEFT JOIN ##Temp_extended_Properties_' + @SessionId + ' as e ON e.major_id = c.object_id and e.minor_id = c.column_id and e.class = 1
				LEFT OUTER JOIN ' + @DBName + '.sys.types AS usrt ON usrt.user_type_id = c.user_type_id
				LEFT OUTER JOIN ' + @DBName + '.sys.types AS baset ON (baset.user_type_id = c.system_type_id and baset.user_type_id = baset.system_type_id) or 
					((baset.system_type_id = c.system_type_id) and (baset.user_type_id = c.user_type_id) and (baset.is_user_defined = 0) and (baset.is_assembly_type = 1)) 
				LEFT OUTER JOIN ' + @DBName + '.sys.computed_columns	as cc ON cc.object_id = c.object_id and c.column_id = cc.column_id
			WHERE o.type in (''U'',''V'')
			ORDER BY c.object_id, c.column_id;
			CREATE CLUSTERED INDEX CLIX_Object_Columns_' + @SessionId + ' ON ##Object_Columns_' + @SessionId + '(object_id, column_id);
			';

    EXEC sp_executesql @TempString

END

/*2110 Looping through tables and views */
WHILE EXISTS (SELECT TOP 1 1 FROM @Objects WHERE ObjectType in ('U','V') and Reported = 0)
BEGIN
  SET @CurrentNumber = @CurrentNumber + 1;
  IF @CurrentNumber > @NextMilestone
  BEGIN 
	SELECT @NextMilestone = @NextMilestone + @IncrementCount;
	SET @CntMessage = 'Processed ' + CAST(CAST(Round(@CurrentNumber*100./@TotalNumber,0) as INT) as varchar) + ' percent of Database objects.'
	RAISERROR (@CntMessage, 0, 1) WITH NOWAIT;
  END

  SET @CurrentStep = '2110';
  BEGIN TRY
    SELECT @object_id = ObjectID, @OType = ObjectType, @OType_Desc = ObjectType_Desc 
    FROM @Objects WHERE ID = ( SELECT MIN(ID) FROM @Objects WHERE ObjectType in ('U','V')  and Reported = 0)
    INSERT INTO @Results(ResultRecord) VALUES 
        ('<TABLE border=0 WIDTH=100%>');
    INSERT INTO @Results(ResultRecord) VALUES 
        ('<TR><TD COLSPAN=2><BR><A id="oid' + CAST(@object_id as VARCHAR)+ '"><HR/></A></TD></TR><TR><TD COLSPAN=2>' +
          (CASE @OType WHEN 'U' THEN '<A HREF="#TableList">Return to list of tables</A>' ELSE '<A HREF="#ViewList">Return to list of Views</A>' END) +
        +'</TD></TR>');
    INSERT INTO @Results(ResultRecord) VALUES 
        ('<TR><TH ALIGN="LEFT" WIDTH=10%>Object ID # </TH><TD>' + CAST(@object_id as varchar) + '</TD></TR>');
    INSERT INTO @Results(ResultRecord) VALUES 
        ('<TR><TH ALIGN="LEFT" >Object Type:</TH><TD>' + @OType_Desc + '</TD></TR>');
    INSERT INTO @Results(ResultRecord) 
    SELECT 
      '<TR><TH ALIGN="LEFT" >Object Schema:</TH><TD>' + IsNull(SchemaName,'&nbsp') + '</TD></TR>'  +
      '<TR><TH ALIGN="LEFT" >Object Name:</TH><TD>' + IsNull(ObjectName,'&nbsp') + '</TD></TR>' +
      '<TR><TH ALIGN="LEFT" >Object Created:</TH><TD>' + IsNull(CONVERT(nvarchar(50),Created_Dt,113),'&nbsp') + '</TD></TR>' + 
      '<TR><TH ALIGN="LEFT" >Object Updated:</TH><TD>' + IsNull(CONVERT(nvarchar(50),Modified_Dt,113),'&nbsp') + '</TD></TR>' + 
      '<TR><TH ALIGN="LEFT" >Object Description:</TH><TD>' + IsNull(Ext_Property,'&nbsp') + '</TD></TR>'
    FROM @Objects WHERE @object_id = ObjectID
  END TRY

  BEGIN CATCH
    SET @ErrorMessage = ERROR_MESSAGE();
    
    BEGIN TRY
        RAISERROR (@ErrorMessage, 16, 0);
    END TRY
    
    BEGIN CATCH
      PRINT @TempString; PRINT ERROR_MESSAGE();
      SELECT @ErrorList = @ErrorList + 
        'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
        'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
    END CATCH
  END CATCH


  INSERT INTO @Results(ResultRecord) VALUES ('</TABLE>');

  /*2120 Going through columns */
  SET @CurrentStep = '2120';
  INSERT INTO @Results(ResultRecord) VALUES 
      ('<H3>Columns:<H3/><TABLE border=1 cellpadding=5>');
  INSERT INTO @Results(ResultRecord) VALUES 
	    ('<TR><TH>ID</TH><TH>Name</TH><TH>Type</TH><TH>User<BR>Defined</TH><TH>Length</TH>' +
	     '<TH>Identity</TH><TH>Nullable</TH><TH>Collation</TH><TH>Description</TH>' +
	     '<TH>Computed</TH><TH>Computed Definition</TH><TH>Persisted</TH></TR>');
/*
  SET @TempString = '
			SELECT ''<TR><TD ALIGN="CENTER">'' + CAST(c.column_id as varchar) + ''</TD><TD>'' + c.name + ''</TD><TD ALIGN="CENTER">'' + 
			CASE usrt.is_user_defined WHEN 1 THEN ''<A HREF="#UDDT'' + CAST(c.user_type_id as VARCHAR) + ''">'' + UPPER(usrt.name) + ''</A>'' ELSE UPPER(usrt.name)  END + 
			''</TD><TD ALIGN="CENTER">'' + CASE usrt.is_user_defined WHEN 0 THEN ''N'' ELSE ''Y'' END + 
			''</TD><TD ALIGN="CENTER">'' + CAST(CAST(CASE WHEN baset.name IN (N''nchar'', N''nvarchar'') AND c.max_length <> -1 THEN c.max_length/2 ELSE c.max_length END AS int) as varchar) + 
			''</TD><TD ALIGN="CENTER">'' + CASE c.is_identity WHEN 0 THEN ''N'' ELSE ''Y'' END + ''</TD><TD ALIGN="CENTER">'' + CASE c.is_nullable WHEN 0 THEN ''N'' ELSE ''Y'' END + 
			''</TD><TD>'' + IsNull(c.Collation_Name COLLATE ' + @Collation + ',''N/A'') + ''</TD><TD>'' + IsNull(e.value,''&nbsp'') +   
			''</TD><TD ALIGN="CENTER">'' + CASE c.is_computed WHEN 1 THEN ''Y'' ELSE ''N'' END +
			''</TD><TD>'' + IsNull(cc.definition,''N/A'') + ''</TD><TD ALIGN="CENTER">'' + CASE cc.is_persisted WHEN 1 THEN ''Y'' WHEN 0 THEN ''N'' ELSE ''N/A'' END + ''</TD></TR>''	
			FROM ' + @DBName + '.sys.columns as c 
				LEFT JOIN ##Temp_extended_Properties_' + @SessionId + ' as e ON e.major_id = c.object_id and e.minor_id = c.column_id and e.class = 1
				LEFT OUTER JOIN ' + @DBName + '.sys.types AS usrt ON usrt.user_type_id = c.user_type_id
				LEFT OUTER JOIN ' + @DBName + '.sys.types AS baset ON (baset.user_type_id = c.system_type_id and baset.user_type_id = baset.system_type_id) or 
					((baset.system_type_id = c.system_type_id) and (baset.user_type_id = c.user_type_id) and (baset.is_user_defined = 0) and (baset.is_assembly_type = 1)) 
				LEFT OUTER JOIN ' + @DBName + '.sys.computed_columns	as cc ON cc.object_id = c.object_id and c.column_id = cc.column_id
			WHERE c.object_id = @object_id ORDER BY c.column_id;';
*/
  SET @TempString = '
			SELECT ''<TR><TD ALIGN="CENTER">'' + CAST(column_id as varchar) + ''</TD><TD>'' + name + ''</TD><TD ALIGN="CENTER">'' + 
			CASE is_user_defined WHEN ''Y'' THEN ''<A HREF="#UDDT'' + CAST(user_type_id as VARCHAR) + ''">'' + Type_Name + ''</A>'' ELSE Type_Name  END + 
			''</TD><TD ALIGN="CENTER">'' + is_user_defined + 
			''</TD><TD ALIGN="CENTER">'' + max_length + 
			''</TD><TD ALIGN="CENTER">'' + is_identity + ''</TD><TD ALIGN="CENTER">'' + is_nullable + 
			''</TD><TD>'' + Collation_Name + ''</TD><TD>'' + ExtValue +   
			''</TD><TD ALIGN="CENTER">'' + is_computed +
			''</TD><TD>'' + definition + ''</TD><TD ALIGN="CENTER">'' + is_persisted + ''</TD></TR>''	
			FROM ##Object_Columns_' + @SessionId + ' WHERE object_id = @object_id ORDER BY column_id;'

  IF @Debug = 1 PRINT @TempString;
  BEGIN TRY
    INSERT INTO @Results(ResultRecord)  
    EXEC sp_executesql @TempString, N'@object_id int', @object_id = @object_id
  END TRY

  BEGIN CATCH
    SET @ErrorMessage = ERROR_MESSAGE();
    
    BEGIN TRY
        RAISERROR (@ErrorMessage, 16, 0);
    END TRY
    
    BEGIN CATCH
      PRINT @TempString; PRINT ERROR_MESSAGE();
      SELECT @ErrorList = @ErrorList + 
        'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
        'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
    END CATCH
  END CATCH

  INSERT INTO @Results(ResultRecord) VALUES ('</TABLE>');



  /*2125 Going through columns' Statistics */
  SET @CurrentStep = '2125';
  SET @TempString = '
      SELECT @i = COUNT(*) FROM ' + @DBName + '.sys.stats st WHERE st.object_id = @object_id' ;

  IF @Debug = 1 PRINT @TempString;
  BEGIN TRY
    EXEC sp_executesql @TempString, N'@object_id int, @i int OUTPUT', @object_id = @object_id, @i = @i OUTPUT 
  END TRY

  BEGIN CATCH
    SET @ErrorMessage = ERROR_MESSAGE();
    
    BEGIN TRY
        RAISERROR (@ErrorMessage, 16, 0);
    END TRY
    
    BEGIN CATCH
      PRINT @TempString; PRINT ERROR_MESSAGE();
      SELECT @ErrorList = @ErrorList + 
        'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
        'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
    END CATCH
  END CATCH  

IF @i > 0
BEGIN
  /* Output statistics*/
  SET @CurrentStep = '2126';
  INSERT INTO @Results(ResultRecord) VALUES 
      ('<H3>Statistics:<H3/><TABLE border=1 cellpadding=5>');
  INSERT INTO @Results(ResultRecord) VALUES 
	    ('<TR><TH>Statistic Name</TH><TH>On Column(s)</TH><TH>Auto<BR>Created</TH>' +
	     '<TH>Auto<BR>Recompute</TH><TH>Last Updated</TH><TH>Filter</TH></TR>');

	SET @TempString = 'USE [' + @DBName + ']; ' +
				'SELECT 
					''<TR><TD>'' + st.name +
					''</TD><TD>'' +
					SUBSTRING((
						SELECT '', '' + c.name FROM sys.stats_columns as sc 
						INNER JOIN sys.columns as c 
							ON sc.object_id = c.object_id and c.column_id = sc.column_id
						WHERE st.object_id = sc.object_id and st.stats_id = sc.stats_id
						ORDER BY sc.stats_column_id
						FOR XML PATH('''')
					),2,8000)  + 	''</TD><TD ALIGN="CENTER">'' +
					CASE st.auto_created WHEN 1 THEN ''Yes'' ELSE ''No'' END +
					''</TD><TD ALIGN="CENTER">'' + 
					CASE st.no_recompute WHEN 0 THEN ''Yes'' ELSE ''No'' END +
					''</TD><TD>'' +IsNull(CONVERT(VARCHAR,STATS_DATE(st.object_id, st.stats_id),120),''N/A'') +
					''</TD><TD ALIGN="CENTER">'' + ' +
					CASE WHEN @ServerRelease >= 100 THEN 'ISNULL(st.filter_definition ,N''N/A'')' ELSE '''N/A''' END +
					' + ''</TD></TR>''	
				FROM sys.stats st WHERE st.object_id = @object_id ORDER BY st.stats_id;';

	  IF @Debug = 1 PRINT @TempString;
	  BEGIN TRY
		INSERT INTO @Results(ResultRecord)  
		EXEC sp_executesql @TempString, N'@object_id int', @object_id = @object_id
	  END TRY

	  BEGIN CATCH
		SET @ErrorMessage = ERROR_MESSAGE();
	    
		BEGIN TRY
			RAISERROR (@ErrorMessage, 16, 0);
		END TRY
	    
		BEGIN CATCH
		  PRINT @TempString; PRINT ERROR_MESSAGE();
		  SELECT @ErrorList = @ErrorList + 
			'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
			'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
		END CATCH
	  END CATCH

	INSERT INTO @Results(ResultRecord) VALUES ('</TABLE>');
END /* Output statistics*/


  /*2130 Going through Indexes */
  SET @CurrentStep = '2130';
  SET @TempString = '
      SELECT @i = COUNT(*) FROM ' + @DBName + '.sys.columns c 
     INNER JOIN ' + @DBName + '.sys.index_columns ic on ic.column_id = c.column_id and ic.object_id = c.object_id 
      WHERE c.object_id = @object_id' ;			

  IF @Debug = 1 PRINT @TempString;
  BEGIN TRY
    EXEC sp_executesql @TempString, N'@object_id int, @i int OUTPUT', @object_id = @object_id, @i = @i OUTPUT 
  END TRY

  BEGIN CATCH
    SET @ErrorMessage = ERROR_MESSAGE();
    
    BEGIN TRY
        RAISERROR (@ErrorMessage, 16, 0);
    END TRY
    
    BEGIN CATCH
      PRINT @TempString; PRINT ERROR_MESSAGE();
      SELECT @ErrorList = @ErrorList + 
        'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
        'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
    END CATCH
  END CATCH
  
  /*2131 Output Indexes */
  IF @i > 0
  BEGIN
    SET @CurrentStep = '2131';
      INSERT INTO @Results(ResultRecord) VALUES 
          ('<H3>Indexes:<H3/><TABLE border=1  cellpadding=5 width=100%>');
      INSERT INTO @Results(ResultRecord) VALUES 
	        ('<TR><TH>Name</TH><TH>Type</TH><TH>Unique</TH><TH>Ignore<BR/>Dups</TH>' +
		      '<TH>Primary<BR/>Key</TH><TH>Unique<BR/>Constr-t</TH><TH>Fill<BR/>factor</TH>' + 
  		    '<TH>Padded</TH><TH>Disabled</TH><TH>Hypoth-<BR/>etical</TH><TH>Row<BR/>Locks</TH>' + 
	  	    '<TH>Page<BR/>Locks</TH><TH>Filter</TH><TH>Max<BR/>Length</TH>' + 
		      '<TH>Indexed<BR/>Columns</TH><TH>Included<BR/>Columns</TH><TH>List of<BR/>Columns</TH>' +
		     '<TH>Index Description</TH></TR>' );

      SET @TempString = '
            ;WITH Totals as (
				            SELECT ic.index_id, Sum(c.max_length) AS MaxLength, 
				              SUM(CASE ic.is_included_column WHEN 0 THEN 1 ELSE 0 END) as Columns_in_Index,
				             SUM(CASE ic.is_included_column WHEN 1 THEN 1 ELSE 0 END) as Included_Columns
				            FROM ' + @DBName + '.sys.columns c 
				            INNER JOIN ' + @DBName + '.sys.index_columns ic on ic.column_id = c.column_id 
				            WHERE ic.object_id = c.object_id and c.object_id = @object_id
				            GROUP BY ic.index_id
            )
		      SELECT 
			      ''<TR><TD>'' + i.name + ''</TD><TD ALIGN="CENTER">'' + 
			      /*CASE i.type WHEN 1 THEN ''Y'' ELSE ''N'' END  + ''</TD><TD ALIGN="CENTER">'' + */
			      i.Type_Desc COLLATE ' + @Collation + ' + ''</TD><TD ALIGN="CENTER">'' + 
			      CASE i.is_unique WHEN 1 THEN ''Y'' ELSE ''N'' END  + ''</TD><TD ALIGN="CENTER">'' + 
			      CASE i.ignore_dup_key WHEN 1 THEN ''Y'' ELSE ''N'' END  + ''</TD><TD ALIGN="CENTER">'' + 
			      CASE i.is_primary_key WHEN 1 THEN ''Y'' ELSE ''N'' END  + ''</TD><TD ALIGN="CENTER">'' + 
			      CASE i.is_unique_constraint WHEN 1 THEN ''Y'' ELSE ''N'' END  + ''</TD><TD ALIGN="CENTER">'' + 
			      CAST(i.fill_factor AS VARCHAR) + ''</TD><TD ALIGN="CENTER">'' + 
			      CASE i.is_padded WHEN 1 THEN ''Y'' ELSE ''N'' END  + ''</TD><TD ALIGN="CENTER">'' + 
			      CASE i.is_disabled WHEN 1 THEN ''Y'' ELSE ''N'' END  + ''</TD><TD ALIGN="CENTER">'' + 
			      CASE i.is_hypothetical WHEN 1 THEN ''Y'' ELSE ''N'' END  + ''</TD><TD ALIGN="CENTER">'' + 
			      CASE i.allow_row_locks WHEN 1 THEN ''OK'' ELSE ''No'' END  + ''</TD><TD ALIGN="CENTER">'' + 
			      CASE i.allow_page_locks WHEN 1 THEN ''OK'' ELSE ''No'' END  + ''</TD><TD ALIGN="CENTER">'' + ' + 
			      
				CASE WHEN @ServerRelease >= 100 THEN 'ISNULL(i.filter_definition ,N''N/A'')' ELSE '''N/A''' END +
				' + ''</TD><TD ALIGN="CENTER">'' + 
			      CAST(t.MaxLength AS VARCHAR) + ''</TD><TD ALIGN="CENTER">'' + 
			      CAST(t.Columns_in_Index AS VARCHAR) + ''</TD><TD ALIGN="CENTER">'' + 
			      CAST(t.Included_Columns AS VARCHAR) + ''</TD><TD>'' + 
			      reverse(stuff(LTRIM(reverse(
			      (
				      SELECT c.Name + ''('' + CASE WHEN ic.is_included_column = 1 then ''I-'' else '''' End + CAST(c.max_length as varchar) + ''), '' 
				      FROM ' + @DBName + '.sys.columns c 
				      INNER JOIN ' + @DBName + '.sys.index_columns ic on ic.column_id = c.column_id 
				      WHERE ic.object_id = i.object_id and c.object_id = i.object_id and ic.index_id = i.index_id 
				      ORDER BY CASE WHEN ic.is_included_column = 1 then 255 else 0 End + ic.key_ordinal
				      for xml path('''')
			      )
			      )), 1, 1, ''''))  + ''</TD><TD>'' + IsNull(CAST(e.value as VARCHAR(MAX)),''&nbsp'') + ''</TD></TR>''  
		      FROM ' + @DBName + '.sys.INDEXES i 
		      INNER JOIN Totals t on t.index_id = i.index_id
 		      LEFT JOIN ##Temp_extended_Properties_' + @SessionId + ' as e ON e.major_id = i.object_id and e.minor_id = i.index_id  and e.class = 7
		      WHERE i.object_id = @object_id and i.name is not null
		      ORDER BY i.index_id, i.name;';
        				
      IF @Debug = 1 PRINT @TempString;
      BEGIN TRY
        INSERT INTO @Results(ResultRecord)
        EXEC sp_executesql @TempString, N'@object_id int', @object_id = @object_id
      END TRY

      BEGIN CATCH
        SET @ErrorMessage = ERROR_MESSAGE();
        
        BEGIN TRY
            RAISERROR (@ErrorMessage, 16, 0);
        END TRY
        
        BEGIN CATCH
          PRINT @TempString; PRINT ERROR_MESSAGE();
          SELECT @ErrorList = @ErrorList + 
            'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
            'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
        END CATCH
      END CATCH

      INSERT INTO @Results(ResultRecord) VALUES ('</TABLE>');
  END
  
  /*2140 Going through constraints */
  SET @CurrentStep = '2140';
  IF EXISTS (SELECT top 1 1  FROM @Objects as o WHERE  o.ParentObjectID = @object_id and o.ObjectType in ('PK', 'D', 'F', 'C', 'UQ') )
  BEGIN
    INSERT INTO @Results(ResultRecord) VALUES (
      '<H3>Constraints:<H3/><TABLE border=1 cellpadding=5>' + 
      '<TR><TH>Name</TH><TH>Type</TH>' + 
      '<TH>Text</TH><TH>Referenced</TH><TH>Description</TH></TR>'
      );
  SET @TempString = '
		  SELECT DISTINCT ''<TR><TD>'' + o.Name COLLATE ' + @Collation + ' + ''</TD><TD>'' +  o.type_desc COLLATE ' + @Collation + ' + ''</TD><TD>'' +  
			  ISNull(CASE o.type
			  WHEN  ''F'' THEN
				  REPLACE((
					  SELECT Name + '','' FROM ' + @DBName + '.sys.columns as co
					  INNER JOIN ' + @DBName + '.sys.foreign_key_columns ifkc1 on ifkc1.parent_column_id = co.column_id and co.object_id = ifkc1.parent_object_id
					  WHERE ifkc1.constraint_object_id = o.object_id
					  ORDER BY co.column_id for xml path('''')
				  ) + ''|'', '',|'','''') + '' -> '' + s.Name + ''.'' + t.Name + ''('' +
				  REPLACE((
					  SELECT Name + '','' FROM ' + @DBName + '.sys.columns as rco
					  INNER JOIN ' + @DBName + '.sys.foreign_key_columns ifkc2 on ifkc2.referenced_column_id = rco.column_id and rco.object_id = ifkc2.referenced_object_id
					  WHERE ifkc2.referenced_object_id = fc.referenced_object_id and ifkc2.parent_object_id = fc.parent_object_id
					  ORDER BY rco.column_id for xml path('''')
				  ) + ''|'', '',|'','''') + '')''
			  WHEN ''PK'' THEN 
				  REPLACE((
					  SELECT c.Name + '','' FROM ' + @DBName + '.sys.columns c 
					  INNER JOIN ' + @DBName + '.sys.index_columns ic on ic.column_id = c.column_id and c.object_id = ic.object_id
					  INNER JOIN ' + @DBName + '.sys.indexes as i on c.object_id = i.object_id and ic.index_id = i.index_id 
					  WHERE c.object_id = o.parent_object_id and i.name = o.name
					  ORDER BY c.column_id for xml path('''')
				  ) + ''|'', '',|'','''') 
			  WHEN ''UQ'' THEN 
				  REPLACE((
					  SELECT c.Name + '','' FROM ' + @DBName + '.sys.columns c 
					  INNER JOIN ' + @DBName + '.sys.index_columns ic on ic.column_id = c.column_id and c.object_id = ic.object_id
					  INNER JOIN ' + @DBName + '.sys.indexes as i on c.object_id = i.object_id and ic.index_id = i.index_id 
					  WHERE c.object_id = o.parent_object_id and i.name = o.name
					  ORDER BY c.column_id for xml path('''')
				  ) + ''|'', '',|'','''') 
			  WHEN ''D'' THEN  c.text + '' -> "'' + dc.name + ''"''
			  ELSE c.text END,''Error!'')  + ''</TD><TD>'' + 
			  CASE WHEN t.name IS NULL THEN ''&nbsp'' ELSE 
			    ''<A HREF="#oid'' + CAST(t.object_id as VARCHAR) + ''">'' + t.name + 
			    CASE t.type WHEN ''U'' THEN '' (Table)'' WHEN ''V'' THEN '' (View)'' ELSE '''' END + ''</A>''
			  END + ''</TD><TD>'' + 
			  IsNull(CAST(e.value as varchar(256)),''&nbsp'')  + ''</TD></TR>''
		  FROM ' + @DBName + '.sys.objects as o
		  LEFT JOIN ' + @DBName + '.sys.syscomments as c on o.object_id = c.id
		  LEFT JOIN ' + @DBName + '.sys.foreign_key_columns as fc ON o.type = ''F'' and fc.constraint_object_id = o.object_id
		  LEFT JOIN ' + @DBName + '.sys.objects as t ON fc.referenced_object_id = t.object_id
		  LEFT JOIN ' + @DBName + '.sys.schemas as s ON s.schema_Id = t.schema_Id
		  LEFT JOIN ##Temp_extended_Properties_' + @SessionId + ' as e ON e.major_id = o.object_id  and e.Minor_ID = 0
		  LEFT JOIN ' + @DBName + '.sys.columns as dc ON dc.object_id = o.parent_object_id and dc.default_object_id = o.object_id
		  WHERE o.type in (''PK'', ''D'', ''F'', ''C'', ''UQ'')
		  and o.parent_object_id = @object_id';

    IF @Debug = 1 PRINT @TempString;
    BEGIN TRY
      INSERT INTO @Results(ResultRecord)  
      EXEC sp_executesql @TempString, N'@object_id int', @object_id = @object_id
    END TRY

    BEGIN CATCH
      SET @ErrorMessage = ERROR_MESSAGE();
      
      BEGIN TRY
          RAISERROR (@ErrorMessage, 16, 0);
      END TRY
      
      BEGIN CATCH
        PRINT @TempString; PRINT ERROR_MESSAGE();
        SELECT @ErrorList = @ErrorList + 
          'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
          'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
      END CATCH
    END CATCH
  		
    INSERT INTO @Results(ResultRecord) VALUES ('</TABLE>');
  END

  /*2150 Going through Dependent Objects */
  SET @CurrentStep = '2150';
  SET @TempString = '
      SELECT top 1 @i = 1 FROM ' + @DBName + '.sys.sql_dependencies as d
      RIGHT JOIN ' + @DBName + '.sys.objects as o on d.object_id = o.object_id
      WHERE o.type not in (''PK'', ''D'', ''F'', ''C'', ''UQ'') 
		        and d.referenced_major_id = @object_id and (o.object_id != @object_id or o.object_id IN 
          ( SELECT parent_object_id FROM ' + @DBName + '.sys.foreign_key_columns WHERE referenced_object_id = @object_id));' ;
  IF @Debug = 1 PRINT @TempString;

  SET @i = 0;
  EXEC sp_executesql @TempString, N'@object_id int, @i int OUTPUT', @object_id = @object_id, @i = @i OUTPUT 
  SET @i = Isnull(@i,0);

  IF @Debug = 1 SELECT @i as "Dependent Objects"

  IF @i > 0
  BEGIN
    INSERT INTO @Results(ResultRecord) VALUES (
      '<H3>Dependent Objects:<H3/><TABLE border=1 cellpadding=5>' + 
      '<TR><TH>##</TH><TH>Schema</TH><TH>Name</TH><TH>Type</TH>' + 
      '<TH>Object ID</TH></TR>');

    SET @TempString = '
    ;WITH DependObj as (
      SELECT DISTINCT o.object_id, sc.Name COLLATE ' + @Collation + ' as Schema_Nm, 
					o.Name, o.type_desc COLLATE ' + @Collation + ' as type_desc
      FROM ' + @DBName + '.sys.sql_dependencies as d
	  INNER JOIN ' + @DBName + '.sys.objects as o on d.object_id = o.object_id
	  INNER JOIN ' + @DBName + '.sys.schemas as sc ON o.schema_id = sc.schema_id
	  WHERE o.type not in (''PK'', ''D'', ''F'', ''C'', ''UQ'') 
		      and d.referenced_major_id = @object_id and o.object_id != @object_id
		  UNION
      SELECT DISTINCT o.object_id, sc.Name COLLATE ' + @Collation + ' as Schema_Nm, 
					o.Name, o.type_desc COLLATE ' + @Collation + ' as type_desc
      FROM ' + @DBName + '.sys.objects as o 
      INNER JOIN ' + @DBName + '.sys.schemas as sc ON o.schema_id = sc.schema_id
	  WHERE o.type not in (''PK'', ''D'', ''F'', ''C'', ''UQ'') and o.parent_object_id = @object_id
		  UNION
	  SELECT DISTINCT o.object_id, sc.Name COLLATE ' + @Collation + ' as Schema_Nm, 
						o.Name, CASE o.type WHEN ''U'' THEN ''Table'' ELSE ''View'' END + '' (via FK)''
      FROM ' + @DBName + '.sys.foreign_key_columns as fkc
      INNER JOIN ' + @DBName + '.sys.objects as o ON o.object_id = fkc.parent_object_id
	  INNER JOIN ' + @DBName + '.sys.schemas as sc ON o.schema_id = sc.schema_id
      WHERE fkc.referenced_object_id = @object_id
     )
		    SELECT ''<TR><TD ALIGN="CENTER">'' +  CAST(ROW_NUMBER() over(order by name, object_id) as VARCHAR) + 
		    ''</TD><TD ALIGN="CENTER">'' + Schema_Nm + 
		    ''</TD><TD><A href="#oid'' + CAST(object_id as VARCHAR) + ''">'' + Name + 
		    ''</TD><TD ALIGN="CENTER">'' + type_desc + 
		    ''</TD><TD ALIGN="CENTER">'' + CAST(object_id as VARCHAR)  + ''</TD></TR>''  
		    FROM DependObj
		   ORDER BY name, object_id ;';

    IF @Debug = 1 PRINT @TempString;
    BEGIN TRY
      INSERT INTO @Results(ResultRecord)  
      EXEC sp_executesql @TempString, N'@object_id int', @object_id = @object_id
    END TRY

    BEGIN CATCH
      SET @ErrorMessage = ERROR_MESSAGE();
      
      BEGIN TRY
          RAISERROR (@ErrorMessage, 16, 0);
      END TRY
      
      BEGIN CATCH
        PRINT @TempString; PRINT ERROR_MESSAGE();
        SELECT @ErrorList = @ErrorList + 
          'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
          'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
      END CATCH
    END CATCH
    
    INSERT INTO @Results(ResultRecord) VALUES ('</TABLE>');
  END

  /*2160 Going through Objects that Object depends on*/
  SET @CurrentStep = '2160';
  IF @OType = 'V' or @OType = 'U'
  BEGIN
  
      SET @TempString = '
        SELECT top 1 @i = 1 FROM ' + @DBName + '.sys.objects as o
		LEFT JOIN ' + @DBName + '.sys.sql_dependencies as d ON d.referenced_major_id = o.object_id and d.object_id = @object_id
		LEFT JOIN ' + @DBName + '.sys.foreign_key_columns as fc ON fc.referenced_object_id = o.object_id and fc.parent_object_id = @object_id
        WHERE d.referenced_major_id is not null or fc.referenced_object_id is not null;' ;

      IF @Debug = 1 PRINT @TempString;
      BEGIN TRY
		SET @i = 0;
        EXEC sp_executesql @TempString, N'@object_id int, @i int OUTPUT', @object_id = @object_id, @i = @i OUTPUT 
		SET @i = Isnull(@i,0);
      END TRY

      BEGIN CATCH
        SET @ErrorMessage = ERROR_MESSAGE();
        
        BEGIN TRY
            RAISERROR (@ErrorMessage, 16, 0);
        END TRY
        
        BEGIN CATCH
          PRINT @TempString; PRINT ERROR_MESSAGE();
          SELECT @ErrorList = @ErrorList + 
            'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
            'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
        END CATCH
      END CATCH

	  
      IF @Debug = 1 SELECT @object_id as "object_id", @i as "Objects that Object depends on" 

      IF @i > 0
      BEGIN
        SET @CurrentStep = '2161';
        INSERT INTO @Results(ResultRecord) VALUES (
          '<H3>That Object Depends On :<H3/><TABLE border=1 cellpadding=5>' + 
          '<TR><TH>##</TH><TH>Schema</TH><TH>Name</TH><TH>Type</TH>' + 
          '<TH>Object ID</TH></TR>');

        SET @TempString = '
        ;WITH DependObj as (
				SELECT DISTINCT o.object_id, sc.Name COLLATE ' + @Collation + ' as Schema_Nm, 
								o.Name, o.type_desc COLLATE ' + @Collation + ' as type_desc
		        FROM ' + @DBName + '.sys.sql_dependencies as d
		        INNER JOIN ' + @DBName + '.sys.objects as o on d.referenced_major_id = o.object_id 
	            INNER JOIN ' + @DBName + '.sys.schemas as sc ON o.schema_id = sc.schema_id
		        WHERE d.object_id = @object_id
		        UNION
				SELECT DISTINCT o.object_id, sc.Name COLLATE ' + @Collation + ' as Schema_Nm, 
								o.Name, CASE o.type WHEN ''U'' THEN ''Table'' ELSE ''View'' END + '' (via FK)''
				FROM ' + @DBName + '.sys.foreign_key_columns as fkc
				INNER JOIN ' + @DBName + '.sys.objects as o ON o.object_id = fkc.referenced_object_id
				INNER JOIN ' + @DBName + '.sys.schemas as sc ON o.schema_id = sc.schema_id
				WHERE fkc.parent_object_id = @object_id
			 )
		        SELECT ''<TR><TD ALIGN="CENTER">'' +  CAST(ROW_NUMBER() over(order by name, object_id) as VARCHAR) + 
		        ''</TD><TD ALIGN="CENTER">'' + Schema_Nm + 
		        ''</TD><TD><A href="#oid'' + CAST(object_id as VARCHAR) + ''">'' + Name + 
		        ''</TD><TD ALIGN="CENTER">'' + type_desc + 
		        ''</TD><TD ALIGN="CENTER">'' + CAST(object_id as VARCHAR)  + ''</TD></TR>''  
		        FROM DependObj
		       ORDER BY name, object_id ;';

        IF @Debug = 1 PRINT @TempString;
        BEGIN TRY
          INSERT INTO @Results(ResultRecord)  
          EXEC sp_executesql @TempString, N'@object_id int', @object_id = @object_id
        END TRY

        BEGIN CATCH
          SET @ErrorMessage = ERROR_MESSAGE();
          
          BEGIN TRY
              RAISERROR (@ErrorMessage, 16, 0);
          END TRY
          
          BEGIN CATCH
            PRINT @TempString; PRINT ERROR_MESSAGE();
            SELECT @ErrorList = @ErrorList + 
              'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
              'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
          END CATCH
        END CATCH

        INSERT INTO @Results(ResultRecord) VALUES ('</TABLE>');
      END
  END

  /*2170 Output View code */
  SET @CurrentStep = '2170';
  IF @OType = 'V'
  BEGIN
      INSERT INTO @Results(ResultRecord) VALUES 
        ('<DIV  id="ObjShow'+ CAST(@object_id as VARCHAR) + '"><H4><a href="javascript:HideShowCode(1,''ObjCode' + CAST(@object_id as VARCHAR) + ''');HideShowCode(0,''ObjShow' + CAST(@object_id as VARCHAR) + ''');">Show Code</a></H4></DIV>');
      INSERT INTO @Results(ResultRecord) VALUES 
        ('<DIV style="display:none;" id="ObjCode' + CAST(@object_id as VARCHAR) + '">');
      INSERT INTO @Results(ResultRecord) VALUES 
        ('<H4><a href="javascript:HideShowCode(0,''ObjCode' + CAST(@object_id as VARCHAR) + ''');HideShowCode(1,''ObjShow' + CAST(@object_id as VARCHAR) + ''');">Hide Code</a></H4>');
      INSERT INTO @Results(ResultRecord) VALUES 
        ('<TABLE border=1 cellpadding=5><TR><TD>');

        SET @TempString = '
          SELECT TOP 1 REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(definition,''<'',''&lt''),''>'',''&gt''), CHAR(13) + CHAR(10), ''<BR>''), CHAR(9), ''&nbsp&nbsp&nbsp&nbsp'') ,CHAR(32), ''&nbsp'')' + 
          ' FROM ' + @DBName + '.sys.sql_modules
          WHERE object_id = @object_id;';
        IF @Debug = 1 PRINT @TempString;
        INSERT INTO @Results(ResultRecord)  
        EXEC sp_executesql @TempString, N'@object_id int', @object_id = @object_id

        INSERT INTO @Results(ResultRecord) VALUES ('</TD></TR></TABLE></DIV>');
  END

  UPDATE @Objects SET Reported = 1 WHERE @object_id = ObjectID;
END

INSERT INTO @Results(ResultRecord) VALUES ('</DIV>');

RAISERROR ('#2110 Finished', 0, 1) WITH NOWAIT;

/********************************************************************************************************************************************************************************************************************/

/*2200 Looping through Stored Procedures, Functions, Triggers */
INSERT INTO @Results(ResultRecord) VALUES
	('<DIV  id="ShowProgrammedObjects"><H2>Programmed Objects'' Details (<a href="javascript:HideShowCode(0,''ShowProgrammedObjects'');HideShowCode(1,''HideProgrammedObjects'');"> SHOW </a>)</H2></DIV>');
INSERT INTO @Results(ResultRecord) VALUES
	('<DIV style="display:none;" id="HideProgrammedObjects"><H2>Programmed Objects'' Details (<a href="javascript:HideShowCode(0,''HideProgrammedObjects'');HideShowCode(1,''ShowProgrammedObjects'');"> HIDE </a>)</H2>')

WHILE EXISTS (SELECT TOP 1 1 FROM @Objects WHERE ObjectType IN ('FN','P','TR','TF') and Reported = 0) 
BEGIN

  SET @CurrentNumber = @CurrentNumber + 1;
  IF @CurrentNumber > @NextMilestone
  BEGIN 
	SELECT @NextMilestone = @NextMilestone + @IncrementCount;
	SET @CntMessage = 'Processed ' + CAST(CAST(Round(@CurrentNumber*100./@TotalNumber,0) as INT) as varchar) + ' percent of Database objects.'
	RAISERROR (@CntMessage, 0, 1) WITH NOWAIT;
  END

  SET @CurrentStep = '2200';
  
  BEGIN TRY
    SELECT @object_id = ObjectID, @OType_Desc = o.ObjectType_Desc, @OType = o.ObjectType, @Obj_Name = o.ObjectName 
    FROM @Objects as o WHERE ID = (
		SELECT /*TOP 1*/ MIN(ID) FROM @Objects WHERE ObjectType IN ('FN','P','TR','TF') and Reported = 0
--		ORDER BY CASE ObjectType WHEN 'P' THEN 1 WHEN 'FN' THEN 2 WHEN 'TF' THEN 3 WHEN 'TR' THEN 4 END, o.ObjectName 
	)

    INSERT INTO @Results(ResultRecord) VALUES 
        ('<TABLE border=0 WIDTH=100%>');
    INSERT INTO @Results(ResultRecord) VALUES 
        ('<TR><TD COLSPAN=2><A id="oid' + CAST(@object_id as VARCHAR)+ '"><HR/></A></TD></TR>');
    INSERT INTO @Results(ResultRecord) VALUES 
        ('<TR><TD COLSPAN=2><A HREF="#' +
        (CASE @OType WHEN 'P' THEN 'Proc' WHEN 'FN' THEN 'Func' WHEN 'TF' THEN 'Func' WHEN 'TR' THEN 'Trig' END) 
       +  'List">Return to list of ' + 
        (CASE @OType WHEN 'P' THEN 'Procedures' WHEN 'FN' THEN 'Functions' WHEN 'TF' THEN 'Functions' WHEN 'TR' THEN 'Triggers' END) 
        + '</A></TD></TR>')
    INSERT INTO @Results(ResultRecord) VALUES 
        ('<TR><TH ALIGN="LEFT" WIDTH=10%>Object ID # </TH><TD>' + CAST(@object_id as varchar) + '</TD></TR>');
    INSERT INTO @Results(ResultRecord) VALUES 
        ('<TR><TH ALIGN="LEFT" >Object Type:</TH><TD>' + @OType_Desc + '</TD></TR>');

    INSERT INTO @Results(ResultRecord) 
    SELECT 
      '<TR><TH ALIGN="LEFT" >Object Schema:</TH><TD>' + IsNull(SchemaName,'&nbsp') + '</TD></TR>'  +
      '<TR><TH ALIGN="LEFT" >Object Name:</TH><TD>' + IsNull(ObjectName,'&nbsp') + '</TD></TR>' +
      '<TR><TH ALIGN="LEFT" >Object Created:</TH><TD>' + IsNull(CONVERT(nvarchar(50),Created_Dt,113),'&nbsp') + '</TD></TR>' + 
      '<TR><TH ALIGN="LEFT" >Object Updated:</TH><TD>' + IsNull(CONVERT(nvarchar(50),Modified_Dt,113),'&nbsp') + '</TD></TR>' + 
      '<TR><TH ALIGN="LEFT" >Object Description:</TH><TD>' + IsNull(Ext_Property,'&nbsp') + '</TD></TR>'
    FROM @Objects WHERE @object_id = ObjectID

    /* Show parent object if applicable */
    IF @OType = 'TR' and EXISTS ( 
      SELECT TOP 1 1 FROM @Objects as o INNER JOIN @Objects as p
      ON o.ParentObjectID = p.ObjectID WHERE @object_id = o.ObjectID)
    INSERT INTO @Results(ResultRecord) 

    SELECT '<TR><TH ALIGN="LEFT" >Parent Object:</TH><TD><A HREF="#oid' + CAST(o.ParentObjectID as Varchar) + '">' + p.ObjectName + 
      CASE p.ObjectType_Desc WHEN 'USER_TABLE' THEN ' (Table)' WHEN 'VIEW' THEN ' (View)' ELSE '' END
      '</A></TD></TR>'
    FROM @Objects as o INNER JOIN @Objects as p ON o.ParentObjectID = p.ObjectID 
    WHERE @object_id = o.ObjectID
  END TRY

  BEGIN CATCH
    SET @ErrorMessage = ERROR_MESSAGE();
    
    BEGIN TRY
        RAISERROR (@ErrorMessage, 16, 0);
    END TRY
    
    BEGIN CATCH
      PRINT @TempString; PRINT ERROR_MESSAGE();
      SELECT @ErrorList = @ErrorList + 
        'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
        'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
    END CATCH
  END CATCH

  INSERT INTO @Results(ResultRecord) VALUES ('</TABLE>');

  /*2210 Object's parameters */
  SET @CurrentStep = '2210';
  SET @TempString = '
    SELECT @i = COUNT(*) FROM ' + @DBName + '.sys.parameters as p
    WHERE p.object_id = @object_id' ;			
  
  IF @Debug = 1 PRINT @TempString;
  BEGIN TRY
    EXEC sp_executesql @TempString, N'@object_id int, @i int OUTPUT', @object_id = @object_id, @i = @i OUTPUT 
  END TRY

  BEGIN CATCH
    SET @ErrorMessage = ERROR_MESSAGE();
    
    BEGIN TRY
        RAISERROR (@ErrorMessage, 16, 0);
    END TRY
    
    BEGIN CATCH
      PRINT @TempString; PRINT ERROR_MESSAGE();
      SELECT @ErrorList = @ErrorList + 
        'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
        'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
    END CATCH
  END CATCH
  
  IF @Debug = 1 SELECT @i as "Objects Parameteers"

  IF @i > 0
  BEGIN
    SET @CurrentStep = '2211';
    INSERT INTO @Results(ResultRecord) VALUES
		  ('<H3>Object''s Parameters:</H3>');
    INSERT INTO @Results(ResultRecord) VALUES
		  ('<TABLE border=1 cellpadding=5>');
    INSERT INTO @Results(ResultRecord) VALUES
		  ('<TR><TH>ID</TH><TH>Name</TH><TH>Type</TH><TH>Max Bytes</TH><TH>Output</TH></TR>');

      SET @TempString = '
			  SELECT ''<TR>'' +
				  ''<TD ALIGN="CENTER">'' + CAST(p.parameter_id as varchar) + ''</TD>'' +
				  ''<TD>'' + CASE p.name WHEN  '''' THEN ''N/A'' ELSE p.name END + ''</TD>'' +
				  ''<TD>'' + t.name + ''</TD>'' +
				  ''<TD ALIGN="CENTER">'' + CASE p.max_length WHEN -1 THEN ''MAX'' ELSE CAST(p.max_length as varchar) END + ''</TD>'' +
				  ''<TD ALIGN="CENTER">'' + CASE p.is_output WHEN 1 THEN ''Y'' ELSE ''N'' END + ''</TD>'' +
				  ''</TR>''
			  FROM ' + @DBName + '.sys.parameters as p INNER JOIN ' + @DBName + '.sys.types as t on t.user_type_id = p.user_type_id
			  WHERE p.object_id = @object_id;';
      
      IF @Debug = 1 PRINT @TempString;
      BEGIN TRY
        INSERT INTO @Results(ResultRecord)  
        EXEC sp_executesql @TempString, N'@object_id int', @object_id = @object_id
      END TRY

      BEGIN CATCH
        SET @ErrorMessage = ERROR_MESSAGE();
        
        BEGIN TRY
            RAISERROR (@ErrorMessage, 16, 0);
        END TRY
        
        BEGIN CATCH
          PRINT @TempString; PRINT ERROR_MESSAGE();
          SELECT @ErrorList = @ErrorList + 
            'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
            'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
        END CATCH
      END CATCH
      
      INSERT INTO @Results(ResultRecord) VALUES ('</TABLE>');
  END

  /*2220 Going through columns */
  SET @CurrentStep = '2220';
  IF @OType = 'TF'
  BEGIN
    INSERT INTO @Results(ResultRecord) VALUES 
        ('<H3>Output Columns:<H3/><TABLE border=1 cellpadding=5>');
    INSERT INTO @Results(ResultRecord) VALUES 
	      ('<TR><TH>ID</TH><TH>Name</TH><TH>Type</TH><TH>Length</TH><TH>Identity</TH><TH>Nullable</TH><TH>Description</TH></TR>');
    SET @TempString = '
			  SELECT ''<TR><TD ALIGN="CENTER">'' + CAST(c.column_id as varchar) + ''</TD><TD>'' + c.name + ''</TD><TD ALIGN="CENTER">'' + usrt.name + 
			  ''</TD><TD ALIGN="CENTER">'' + CAST(CAST(CASE WHEN baset.name IN (N''nchar'', N''nvarchar'') AND c.max_length <> -1 THEN c.max_length/2 ELSE c.max_length END AS int) as varchar) + 
			  ''</TD><TD ALIGN="CENTER">'' + CASE c.is_identity WHEN 0 THEN ''N'' ELSE ''Y'' END + ''</TD><TD ALIGN="CENTER">'' + CASE c.is_nullable WHEN 0 THEN ''N'' ELSE ''Y'' END + 
			  ''</TD><TD>'' + IsNull(CAST(e.value as varchar(256)),''&nbsp'') + ''</TD></TR>''	
			  FROM ' + @DBName + '.sys.columns as c 
				  LEFT JOIN ' + @DBName + '.sys.extended_properties as e ON e.major_id = c.object_id and e.minor_id = c.column_id and e.class = 1
				  LEFT OUTER JOIN ' + @DBName + '.sys.types AS usrt ON usrt.user_type_id = c.user_type_id
				  LEFT OUTER JOIN ' + @DBName + '.sys.types AS baset ON (baset.user_type_id = c.system_type_id and baset.user_type_id = baset.system_type_id) or 
					  ((baset.system_type_id = c.system_type_id) and (baset.user_type_id = c.user_type_id) and (baset.is_user_defined = 0) and (baset.is_assembly_type = 1)) 
			  WHERE c.object_id = ' + CAST (@object_id as Varchar) + ' ORDER BY c.column_id;';

    IF @Debug = 1 PRINT @TempString;
    BEGIN TRY
      INSERT INTO @Results(ResultRecord)  EXECUTE (@TempString);
    END TRY

    BEGIN CATCH
      SET @ErrorMessage = ERROR_MESSAGE();
      
      BEGIN TRY
          RAISERROR (@ErrorMessage, 16, 0);
      END TRY
      
      BEGIN CATCH
        PRINT @TempString; PRINT ERROR_MESSAGE();
        SELECT @ErrorList = @ErrorList + 
          'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
          'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
      END CATCH
    END CATCH
    INSERT INTO @Results(ResultRecord) VALUES ('</TABLE>');
  END

  /*2230 Going through Dependent Objects */
  SET @CurrentStep = '2220';
  SET @TempString = '
      SELECT top 1 @i = 1 FROM ' + @DBName + '.sys.sql_dependencies as d
      INNER JOIN ' + @DBName + '.sys.objects as o on d.object_id = o.object_id
      WHERE o.type not in (''PK'', ''D'', ''F'', ''C'', ''UQ'') 
		        and d.referenced_major_id = @object_id and o.object_id != @object_id' ;			
  IF @Debug = 1 PRINT @TempString;
  EXEC sp_executesql @TempString, N'@object_id int, @i int OUTPUT', @object_id = @object_id, @i = @i OUTPUT 
  SET @i = Isnull(@i,0);

  IF @Debug = 1 SELECT @i as "Dependent Objects"

  IF @i > 0
  BEGIN
    SET @CurrentStep = '2221';
    INSERT INTO @Results(ResultRecord) VALUES (
      '<H3>Dependent Objects:<H3/><TABLE border=1 cellpadding=5>' + 
      '<TR><TH>##</TH><TH>Schema</TH><TH>Name</TH><TH>Type</TH>' + 
      '<TH>Object ID</TH></TR>');

    SET @TempString = '
    ;WITH DependObj as (
			SELECT DISTINCT o.object_id,  sc.Name COLLATE ' + @Collation + ' as Schema_Nm, 
					o.Name, o.type_desc COLLATE ' + @Collation + ' as type_desc
		    FROM ' + @DBName + '.sys.sql_dependencies as d
		    INNER JOIN ' + @DBName + '.sys.objects as o on d.object_id = o.object_id
			INNER JOIN ' + @DBName + '.sys.schemas as sc ON o.schema_id = sc.schema_id
		    WHERE /*o.type not in (''PK'', ''D'', ''F'', ''C'', ''UQ'') 
		      and*/ d.referenced_major_id = @object_id and o.object_id != @object_id
     )
		    SELECT ''<TR><TD ALIGN="CENTER">'' +  CAST(ROW_NUMBER() over(order by name, object_id) as VARCHAR) + 
		    ''</TD><TD ALIGN="CENTER">'' + Schema_Nm +
		    ''</TD><TD><A href="#oid'' + CAST(object_id as VARCHAR) + ''">'' + Name + 
		    ''</TD><TD ALIGN="CENTER">'' + type_desc + 
		    ''</TD><TD ALIGN="CENTER">'' + CAST(object_id as VARCHAR)  + ''</TD></TR>''  
		    FROM DependObj
		   ORDER BY name, object_id ;';

    IF @Debug = 1 PRINT @TempString;
    BEGIN TRY
      INSERT INTO @Results(ResultRecord)  
      EXEC sp_executesql @TempString, N'@object_id int', @object_id = @object_id
    END TRY

    BEGIN CATCH
      SET @ErrorMessage = ERROR_MESSAGE();
      
      BEGIN TRY
          RAISERROR (@ErrorMessage, 16, 0);
      END TRY
      
      BEGIN CATCH
        PRINT @TempString; PRINT ERROR_MESSAGE();
        SELECT @ErrorList = @ErrorList + 
          'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
          'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
      END CATCH
    END CATCH

    INSERT INTO @Results(ResultRecord) VALUES ('</TABLE>');
  END;

  /*2235 Going through Possibly Dependent Objects (forFunctions only) */
  SET @CurrentStep = '2235';
  IF @OType = 'FN' and @CheckFunctionDependents = 1
  BEGIN
      DELETE FROM  @Dependent; 
      
      SET @TempString = '
          SELECT m.object_id, s.Name, o.name, o.type_desc 
          FROM ' + @DBName + '.sys.sql_modules as m
          INNER JOIN ' + @DBName + '.sys.objects as o ON m.object_id = o.object_id
          INNER JOIN ' + @DBName + '.sys.schemas as s ON s.schema_id = o.schema_id 
          WHERE definition like ''%'' + @Obj_Name + ''%'' and m.object_id != @Object_id
              and NOT EXISTS (SELECT TOP 1 1 FROM ' + @DBName + '.sys.sql_dependencies as d WHERE d.object_id = @Object_id and d.referenced_major_id = m.object_id)
          ORDER BY s.Name, o.name;' ;			
      IF @Debug = 1 PRINT @TempString + '              ' + CAST(@object_id as VARCHAR);

      BEGIN TRY
        INSERT INTO @Dependent ([Object_id], [Schema_Name], [Object_Name], ObjectTypeDesc)
        EXEC sp_executesql @TempString, N'@object_id int, @Obj_Name VARCHAR(128)', @object_id = @object_id, @Obj_Name = @Obj_Name
      END TRY

      BEGIN CATCH
        SET @ErrorMessage = ERROR_MESSAGE();
        
        BEGIN TRY
            RAISERROR (@ErrorMessage, 16, 0);
        END TRY
        
        BEGIN CATCH
          PRINT @TempString; PRINT ERROR_MESSAGE();
          SELECT @ErrorList = @ErrorList + 
            'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
            'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
        END CATCH
      END CATCH

      IF @Debug = 1 SELECT * FROM @Dependent;

      IF EXISTS (SELECT top 1 1 FROM @Dependent ) 
      BEGIN
        BEGIN TRY
          INSERT INTO @Results(ResultRecord) VALUES (
            '<H3>Possibly Dependent Objects:<H3/><TABLE border=1 cellpadding=5>' + 
            '<TR><TH>##</TH><TH>Name</TH><TH>Type</TH>' + 
            '<TH>Object ID</TH></TR>');

          INSERT INTO @Results(ResultRecord)  
          SELECT '<TR><TD ALIGN="CENTER">' + CAST(ID as VARCHAR) + 
              '</TD><TD>' + [Schema_Name] +
              '</TD><TD><A href="#oid' + CAST([Object_id] as VARCHAR) + '">' + [Object_Name] + 
              '</TD><TD>' + ObjectTypeDesc + '</TD></TR>'
          FROM @Dependent
          ORDER BY [Object_id];
          DELETE FROM @Dependent;
        END TRY

        BEGIN CATCH
          SET @ErrorMessage = ERROR_MESSAGE();
          
          BEGIN TRY
              RAISERROR (@ErrorMessage, 16, 0);
          END TRY
          
          BEGIN CATCH
            PRINT @TempString; PRINT ERROR_MESSAGE();
            SELECT @ErrorList = @ErrorList + 
              'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
              'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
          END CATCH
        END CATCH

        INSERT INTO @Results(ResultRecord) VALUES ('</TABLE>');
      END;
  END; /* Check for Possibly Dependent Objects (forFunctions only) */

  /*2240 Going through Objects that Object depends on*/
  SET @CurrentStep = '2240';
  SET @TempString = '
    SELECT top 1 @i = 1 FROM ' + @DBName + '.sys.sql_dependencies as d
    INNER JOIN ' + @DBName + '.sys.objects as o on d.referenced_major_id = o.object_id
    WHERE d.object_id = @object_id' ;			

  IF @Debug = 1 PRINT @TempString;
  BEGIN TRY
    EXEC sp_executesql @TempString, N'@object_id int, @i int OUTPUT', @object_id = @object_id, @i = @i OUTPUT 
	SET @i = Isnull(@i,0);
  END TRY


  BEGIN CATCH
    SET @ErrorMessage = ERROR_MESSAGE();
    
    BEGIN TRY
        RAISERROR (@ErrorMessage, 16, 0);
    END TRY
    
    BEGIN CATCH
      PRINT @TempString; PRINT ERROR_MESSAGE();
      SELECT @ErrorList = @ErrorList + 
        'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
        'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
    END CATCH
  END CATCH

  IF @Debug = 1 SELECT @i as "Objects that Object depends on"

  IF @i > 0
  BEGIN
    SET @CurrentStep = '2241';
    INSERT INTO @Results(ResultRecord) VALUES (
      '<H3>That Object Depends On :<H3/><TABLE border=1 cellpadding=5>' + 
      '<TR><TH>##</TH><TH>Schema</TH><THName</TH><TH>Type</TH>' + 
      '<TH>Object ID</TH></TR>');

    SET @TempString = '
    ;WITH DependObj as (
			SELECT DISTINCT o.object_id, sc.Name COLLATE ' + @Collation + ' as Schema_Nm, 
							o.Name, o.type_desc COLLATE ' + @Collation + ' as type_desc
		    FROM ' + @DBName + '.sys.sql_dependencies as d
		    INNER JOIN ' + @DBName + '.sys.objects as o on d.referenced_major_id = o.object_id 
			INNER JOIN ' + @DBName + '.sys.schemas as sc ON o.schema_id = sc.schema_id
		    WHERE d.object_id = @object_id
     )
		    SELECT ''<TR><TD ALIGN="CENTER">'' +  CAST(ROW_NUMBER() over(order by name, object_id) as VARCHAR) + 
		    ''</TD><TD ALIGN="CENTER">'' + Schema_Nm +
		    ''</TD><TD><A href="#oid'' + CAST(object_id as VARCHAR) + ''">'' + Name + 
		    ''</TD><TD ALIGN="CENTER">'' + type_desc + 
		    ''</TD><TD ALIGN="CENTER">'' + CAST(object_id as VARCHAR)  + ''</TD></TR>''  
		    FROM DependObj
		   ORDER BY name, object_id ;';

    IF @Debug = 1 PRINT @TempString;
    BEGIN TRY
      INSERT INTO @Results(ResultRecord)  
      EXEC sp_executesql @TempString, N'@object_id int', @object_id = @object_id
    END TRY

    BEGIN CATCH
      SET @ErrorMessage = ERROR_MESSAGE();
      
      BEGIN TRY
          RAISERROR (@ErrorMessage, 16, 0);
      END TRY
      
      BEGIN CATCH
        PRINT @TempString; PRINT ERROR_MESSAGE();
        SELECT @ErrorList = @ErrorList + 
          'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
          'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
      END CATCH
    END CATCH

    INSERT INTO @Results(ResultRecord) VALUES ('</TABLE>');
  END

  /*2250 Output code */
  SET @CurrentStep = '2250';
  INSERT INTO @Results(ResultRecord) VALUES 
    ('<DIV  id="ObjShow'+ CAST(@object_id as VARCHAR) + '"><H4><a href="javascript:HideShowCode(1,''ObjCode' + CAST(@object_id as VARCHAR) + ''');HideShowCode(0,''ObjShow' + CAST(@object_id as VARCHAR) + ''');">Show Code</a></H4></DIV>');
  INSERT INTO @Results(ResultRecord) VALUES 
    ('<DIV style="display:none;" id="ObjCode' + CAST(@object_id as VARCHAR) + '">');
  INSERT INTO @Results(ResultRecord) VALUES 
    ('<H4><a href="javascript:HideShowCode(0,''ObjCode' + CAST(@object_id as VARCHAR) + ''');HideShowCode(1,''ObjShow' + CAST(@object_id as VARCHAR) + ''');">Hide Code</a></H4>');
  INSERT INTO @Results(ResultRecord) VALUES 
    ('<TABLE border=1 cellpadding=5><TR><TD>');

    SET @TempString = '
      SELECT TOP 1 REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(definition,''<'',''&lt''),''>'',''&gt''), CHAR(13) + CHAR(10), ''<BR>''), CHAR(9), ''&nbsp&nbsp&nbsp&nbsp'') ,CHAR(32), ''&nbsp'')' + 
      ' FROM ' + @DBName + '.sys.sql_modules
      WHERE object_id = @object_id;';

    IF @Debug = 1 PRINT @TempString;
    BEGIN TRY
      INSERT INTO @Results(ResultRecord)  
      EXEC sp_executesql @TempString, N'@object_id int', @object_id = @object_id
    END TRY

    BEGIN CATCH
      SET @ErrorMessage = ERROR_MESSAGE();
      
      BEGIN TRY
          RAISERROR (@ErrorMessage, 16, 0);
      END TRY
      
      BEGIN CATCH
        PRINT @TempString; PRINT ERROR_MESSAGE();
        SELECT @ErrorList = @ErrorList + 
          'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
          'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
      END CATCH
    END CATCH

    INSERT INTO @Results(ResultRecord) VALUES ('</TD></TR></TABLE></DIV>');

    UPDATE @Objects SET Reported = 1 WHERE @object_id = ObjectID;
END

RAISERROR ('Processed 100 percent of Database objects.', 0, 1) WITH NOWAIT;
RAISERROR ('#2200 Finished', 0, 1) WITH NOWAIT;

/********************************************************************************************************************************************************************************************************************/
/*2300 Going throug list of User Defined Data Types if any*/
SET @CurrentStep = '2300';
WHILE EXISTS (SELECT TOP 1 1 FROM @Objects WHERE ObjectType = 'UD' and Reported = 0) 
BEGIN
    SELECT @object_id = ParentObjectID, @OType_Desc = o.ObjectType_Desc
    FROM @Objects as o WHERE ID = ( SELECT MIN(ID) FROM @Objects WHERE ObjectType = 'UD' and Reported = 0)

    /*2310  Insert User Defined Data Type header */
    SET @CurrentStep = '2310';
    INSERT INTO @Results(ResultRecord) VALUES 
        ('<TABLE border=0 WIDTH=100%>');
    INSERT INTO @Results(ResultRecord) VALUES 
        ('<TR><TD COLSPAN=2><A id="UDDT' + CAST(@object_id as VARCHAR)+ '"><HR/></A></TD></TR>');
    INSERT INTO @Results(ResultRecord) VALUES 
        ('<TR><TD COLSPAN=2><A HREF="#UDDTList">Return to list User Defined Data Types</A></TD></TR>');
    INSERT INTO @Results(ResultRecord) VALUES 
        ('<TR><TH ALIGN="LEFT" WIDTH=10%>User Type ID:</TH><TD>' + CAST(@object_id as varchar) + '</TD></TR>');
    INSERT INTO @Results(ResultRecord) VALUES 
        ('<TR><TH ALIGN="LEFT" >Object Type:</TH><TD>' + @OType_Desc + '</TD></TR>');

    INSERT INTO @Results(ResultRecord) 
    SELECT 
      '<TR><TH ALIGN="LEFT" >Object Schema:</TH><TD>' + IsNull(SchemaName,'&nbsp') + '</TD></TR>'  +
      '<TR><TH ALIGN="LEFT" >Object Name:</TH><TD>' + IsNull(ObjectName,'&nbsp') + '</TD></TR>' 
    FROM @Objects WHERE @object_id = ParentObjectID and ObjectType = 'UD';

    SET @TempString = '
        SELECT 
          ''<TR><TH ALIGN="LEFT">Max Length:</TD><TD>'' + CAST(u.max_length/(CASE WHEN t.name IN (''NVARCHAR'',''NCHAR'') THEN 2 ELSE 1 END) as VARCHAR)  + ''</TD></TR>'' +
          ''<TR><TH ALIGN="LEFT">Precision:</TD><TD>'' + CAST(u.Precision as VARCHAR)  + ''</TD></TR>'' +
          ''<TR><TH ALIGN="LEFT">Scale:</TD><TD>'' + CAST(u.scale as VARCHAR)  + ''</TD></TR>'' +
          ''<TR><TH ALIGN="LEFT">Collation Name:</TD><TD>'' + IsNull(u.Collation_name,''&nbsp'')  + ''</TD></TR>'' +
          ''<TR><TH ALIGN="LEFT">Nullable:</TD><TD>'' + CASE u.is_nullable WHEN 1 THEN ''Yes'' ELSE ''No'' END + ''</TD></TR>'' +
          ''<TR><TH ALIGN="LEFT">User Description:</TD><TD>'' + IsNull(CAST(e.value AS VARCHAR(MAX)),''-- Not specified --'')  + ''</TD></TR>'' +
          ''<TR><TH ALIGN="LEFT">System Type ID:</TD><TD>'' + CAST(u.system_type_id as VARCHAR)  + ''</TD></TR>'' +
          ''<TR><TH ALIGN="LEFT">Parent Type:</TD><TD>'' + UPPER (t.name COLLATE ' + @Collation + ') + ''</TD></TR>''
        FROM ' + @DBName + '.sys.types as u 
        INNER JOIN ' + @DBName + '.sys.types as t ON t.user_type_id = u.system_type_id and u.is_user_defined = 1
        INNER JOIN ' + @DBName + '.sys.schemas as s ON u.schema_id = s.schema_id
        LEFT JOIN ##Temp_extended_Properties_' + @SessionId + ' as e ON u.user_type_id = e.major_id and e.minor_id = 0 and e.class = 6
        WHERE u.user_type_id = @uid';

    IF @Debug = 1 PRINT @TempString;
    BEGIN TRY
      INSERT INTO @Results(ResultRecord) EXEC sp_executesql @TempString, N'@uid int', @uid = @object_id
    END TRY

    BEGIN CATCH
      SET @ErrorMessage = ERROR_MESSAGE();
      
      BEGIN TRY
          RAISERROR (@ErrorMessage, 16, 0);
      END TRY
      
      BEGIN CATCH
        PRINT @TempString; PRINT ERROR_MESSAGE();
        SELECT @ErrorList = @ErrorList + 
          'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
          'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
      END CATCH
    END CATCH
    
    INSERT INTO @Results(ResultRecord) VALUES ('</TABLE>');

    /*2320  check if User Defined Data Type is used */
    SET @CurrentStep = '2320';
    SET @TempString = 'SELECT @i = COUNT(*) FROM ' + @DBName + '.sys.columns as c WHERE c.user_type_id = @utid;'; 
    EXEC sp_executesql @TempString, N'@utid int, @i int OUTPUT', @utid = @object_id, @i = @i OUTPUT 
  
    IF @i > 0
    BEGIN
      INSERT INTO @Results(ResultRecord) VALUES (
          '<H3>Listo of Usage of User Defined Data Type:<H3/><TABLE border=1 cellpadding=5>' + 
          '<TR><TH>Object Schema</TH><TH>Object Name</TH><TH>Object Type</TH><TH>Field Name</TH><TH>Nullable</TH><TH>Description</TH></TR>');
      
      /*2330  if User Defined Data Type is used than list all using objects */
      SET @CurrentStep = '2330';
      SET @TempString = '
          SELECT ''<TR><TD>'' + s.name + ''</TD><TD><A HREF="#oid'' + CAST(c.object_id AS VARCHAR) + ''">'' + o.name + ''</A></TD><TD ALIGN="CENTER">'' + 
            o.type_desc COLLATE ' + @Collation + ' + ''</TD><TD>'' + c.name + ''</TD><TD ALIGN="CENTER">'' + 
            CASE c.is_nullable WHEN 0 THEN ''No'' ELSE ''Yes'' END + ''</TD><TD>'' +  IsNull(CAST(e.value AS VARCHAR(MAX)),''&nbsp'')
          FROM ' + @DBName + '.sys.columns as c
          INNER JOIN ' + @DBName + '.sys.objects as o ON c.object_id = o.object_id 
          INNER JOIN ' + @DBName + '.sys.schemas as s ON s.schema_id = o.schema_id
          LEFT JOIN ##Temp_extended_Properties_' + @SessionId + ' as e ON o.object_id  = e.major_id and e.minor_id = c.column_id and e.class = 1
          WHERE c.user_type_id = @utid ORDER BY s.name,  o.name;'; 

      IF @Debug = 1 PRINT @TempString;
      BEGIN TRY
        INSERT INTO @Results(ResultRecord) 
        EXEC sp_executesql @TempString, N'@utid int', @utid = @object_id
      END TRY

      BEGIN CATCH
        SET @ErrorMessage = ERROR_MESSAGE();
        
        BEGIN TRY
            RAISERROR (@ErrorMessage, 16, 0);
        END TRY
        
        BEGIN CATCH
          PRINT @TempString; PRINT ERROR_MESSAGE();
          SELECT @ErrorList = @ErrorList + 
            'ErrorLine: ' + CAST(ERROR_LINE ()-8 as VARCHAR) + ' (#' + @CurrentStep + ')<BR>' +
            'ErrorMessage: "' + ERROR_MESSAGE() + '"<BR><BR>';
        END CATCH
      END CATCH

      INSERT INTO @Results(ResultRecord) VALUES ('</TABLE>');
    END
    
    UPDATE @Objects SET Reported = 1 WHERE @object_id = ParentObjectID and ObjectType = 'UD';
END

INSERT INTO @Results(ResultRecord) VALUES ('</DIV>')

RAISERROR ('#2300 Finished', 0, 1) WITH NOWAIT;

End /* IF @ReportOnlyObjectNames = 0 */

INSERT INTO @Results(ResultRecord) VALUES ('<BR><CENTER><A HREF="http://slavasql.blogspot.com/2013/11/stored-procedure-to-document-database.html">usp_Documenting_DB</CENTER><BR>');
INSERT INTO @Results(ResultRecord) VALUES ('</BODY></HTML>');

IF @ErrorList != ''
  UPDATE @Results
  SET ResultRecord = '<FONT COLOR="RED">' + @ErrorList + '</FONT>'
  WHERE ID = 1;

IF @Debug = 1 SELECT * FROM @Results ORDER BY ID
ELSE SELECT ResultRecord FROM @Results ORDER BY ID

/* Cleanup: Deleting temporary table with Extended_Properties */
SET @TempString = 'DROP TABLE [##Temp_extended_Properties_' + @SessionId + '];'
IF @Debug = 1 PRINT @TempString;
EXECUTE (@TempString);

END_of_SCRIPT:

-- RETURN 0;
------------------------------------------------------------------------------------------------------------------------------------
GO
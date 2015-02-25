/*
View usage space in database
http://www.mssqltips.com/sqlservertip/1805/different-ways-to-determine-free-space-for-sql-server-databases-and-database-files/

*/

-- General
use ILS;
exec sp_helpdb 'ILS';

-- Free space in SQL Server
use ILS;
exec sp_spaceused

-- DBCC SQLPERF to check free space for a SQL Server database
USE master;
DBCC SQLPERF(logspace) 

-- BCC SRHINKFILE to check free space for a SQL database
USE Test5;
DBCC SHRINKFILE (test5) 
DBCC SHRINKFILE (test5_log) 

-- Using FILEPROPERTY to check for free space in a database
use ILS;
SELECT DB_NAME() AS DbName, 
name AS FileName, 
size/128.0 AS CurrentSizeMB,  
size/128.0 - CAST(FILEPROPERTY(name, 'SpaceUsed') AS INT)/128.0 AS FreeSpaceMB 
FROM sys.database_files; 


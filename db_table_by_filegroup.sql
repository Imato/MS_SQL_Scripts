/* View file group of table 
http://www.jasonstrate.com/2013/01/determining-file-group-for-a-table/
*/

-- Listing 1. Query to determine table filegroup by index
SELECT OBJECT_SCHEMA_NAME(t.object_id) AS schema_name, t.name AS table_name,
	i.index_id, i.name AS index_name, ds.name AS filegroup_name, p.rows
FROM sys.tables t
INNER JOIN sys.indexes i ON t.object_id=i.object_id
INNER JOIN sys.filegroups ds ON i.data_space_id=ds.data_space_id
INNER JOIN sys.partitions p ON i.object_id=p.object_id AND i.index_id=p.index_id
ORDER BY t.name, i.index_id

-- Listing 2. Query to determine table filegroup by index and partition
SELECT OBJECT_SCHEMA_NAME(t.object_id) AS schema_name, 
	t.name AS table_name, i.index_id, i.name AS index_name,
	p.partition_number, fg.name AS filegroup_name, p.rows
FROM sys.tables t
INNER JOIN sys.indexes i ON t.object_id = i.object_id
INNER JOIN sys.partitions p ON i.object_id=p.object_id AND i.index_id=p.index_id
LEFT OUTER JOIN sys.partition_schemes ps ON i.data_space_id=ps.data_space_id
LEFT OUTER JOIN sys.destination_data_spaces dds ON ps.data_space_id=dds.partition_scheme_id AND p.partition_number=dds.destination_id
INNER JOIN sys.filegroups fg ON COALESCE(dds.data_space_id, i.data_space_id)=fg.data_space_id
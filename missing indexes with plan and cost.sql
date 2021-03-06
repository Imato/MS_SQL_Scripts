--Missing Indexes Correlated Against Execution Costs and Counts
--http://sqlmag.com/blog/missing-indexes-correlated-against-execution-costs-and-counts

-- based off of a combination of:
--      Jason Strate's Excellent Missing Indexes queries to get the Execution Plans:
--              http://www.jasonstrate.com/2010/12/can-you-dig-it-missing-indexes/
--        And my own technique of grabbing the costs and multiplying them 
--                      by the Execution Count to get Aggregate costs:
--              http://sqlmag.com/blog/performance-tip-find-your-most-expensive-queries
--      TODO: need to figure out a way to 
--              a) get a list of indexes by impact, then 
--              b) cross-apply their plans and get costs. 

SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED;
GO

WITH XMLNAMESPACES(DEFAULT 'http://schemas.microsoft.com/sqlserver/2004/07/showplan'),
PlanMissingIndexes AS (
        SELECT 
                query_plan, 
                usecounts 
        FROM 
                sys.dm_exec_cached_plans cp
                CROSS APPLY sys.dm_exec_query_plan(cp.plan_handle) qp
        WHERE 
                qp.query_plan.exist('//MissingIndexes') = 1
), 
MissingIndexes AS (
        SELECT
                stmt_xml.value('(QueryPlan/MissingIndexes/MissingIndexGroup/MissingIndex/@Database)[1]'
                        , 'sysname') AS DatabaseName,
                stmt_xml.value('(QueryPlan/MissingIndexes/MissingIndexGroup/MissingIndex/@Schema)[1]'
                        , 'sysname') AS SchemaName,
                stmt_xml.value('(QueryPlan/MissingIndexes/MissingIndexGroup/MissingIndex/@Table)[1]'
                        , 'sysname') AS TableName,
                stmt_xml.value('(QueryPlan/MissingIndexes/MissingIndexGroup/@Impact)[1]'
                        , 'float') AS Impact,
                ISNULL(CAST(stmt_xml.value('(@StatementSubTreeCost)[1]'
                        , 'VARCHAR(128)') as float),0) AS Cost,
                pmi.usecounts UseCounts,
                STUFF((SELECT DISTINCT ', ' + c.value('(@Name)[1]', 'sysname')
        FROM 
                stmt_xml.nodes('//ColumnGroup') AS t(cg)
                CROSS APPLY cg.nodes('Column') AS r(c)
        WHERE 
                cg.value('(@Usage)[1]', 'sysname') = 'EQUALITY' 
                FOR  XML PATH('')), 1, 2, '') AS equality_columns
                        ,STUFF((SELECT DISTINCT ', ' + c.value('(@Name)[1]', 'sysname')
                FROM stmt_xml.nodes('//ColumnGroup') AS t(cg)
                CROSS APPLY cg.nodes('Column') AS r(c)
                WHERE cg.value('(@Usage)[1]', 'sysname') = 'INEQUALITY'
                FOR  XML PATH('')), 1, 2, '') AS inequality_columns
                        ,STUFF((SELECT DISTINCT ', ' + c.value('(@Name)[1]', 'sysname')
                FROM stmt_xml.nodes('//ColumnGroup') AS t(cg)
                CROSS APPLY cg.nodes('Column') AS r(c)
                WHERE cg.value('(@Usage)[1]', 'sysname') = 'INCLUDE'
                FOR  XML PATH('')), 1, 2, '') AS include_columns
                ,query_plan
                ,stmt_xml.value('(@StatementText)[1]', 'varchar(4000)') AS sql_text
                FROM PlanMissingIndexes pmi
                CROSS APPLY query_plan.nodes('//StmtSimple') AS stmt(stmt_xml)
                WHERE stmt_xml.exist('QueryPlan/MissingIndexes') = 1
)

SELECT TOP 200
        DatabaseName,
        SchemaName,
        TableName,
        equality_columns,
        inequality_columns,
        include_columns,
        usecounts,
        Cost,
        Cost * UseCounts [AggregateCost],
        Impact,
        query_plan
        FROM MissingIndexes
        ORDER BY 
                Cost * usecounts DESC;


declare	@fiscalNom int = <<>>;
	
set nocount on;

	
IF OBJECT_ID('tempdb..#R_metric_queries') is not null begin drop table #R_metric_queries end;  
if object_id('tempdb..#R_allowedMetrics') is not null begin drop table #R_allowedMetrics end;
/* Allowed Metrics are Metrics that exist in the Weights Table */


/* Grabbing a count of Rows that have the weights for the fiscal Month */
declare @weightRowCount as int;
set @weightRowCount = (
select count(*) from GVPOperations.VID.R_Weights w with(nolock)
where w.fiscalMth = @fiscalNom);

/* If We do not have any weights for that fiscal Month, Copy from the previous */
if @weightRowCount = 0 begin

	insert into GVPOperations.VID.R_Weights(fiscalMth,groupID, locID, metricID, metricWeight, posLevel)
	select 
		@fiscalNom fiscalMth
		,w.groupID
		,w.locID
		,w.metricID
		,w.metricWeight
		,w.posLevel
	from GVPOperations.VID.R_Weights w with(nolock)
	where w.fiscalMth = (@fiscalNom - 1)

end;




select  	
	w.metricID 	
	,w.posLevel 
into #R_allowedMetrics
from GVPOperations.VID.R_Weights w with(nolock) 
where w.fiscalMth = @fiscalNom 
group by  	w.metricID 	,w.posLevel 


/* Metric 32 and 28 Staffed and Calls Handled should be allowed for each posLevel */

insert into #R_allowedMetrics 
select
	32 metricID
	,2 posLevel

insert into #R_allowedMetrics
select
	28 metricID
	,2 posLevel

insert into #R_allowedMetrics
select
	32 metricID
	,3 posLevel

insert into #R_allowedMetrics
select
	28 metricID
	,3 posLevel

insert into #R_allowedMetrics
select
	32 metricID
	,4 posLevel


/* Creating a Table of Metric Insert Queries to Iterate over */
select  
	cast(step as int) step, 
	task, 
	query, 
	cast(am.posLevel as int) posLevel,
	row_number() over(order by am.posLevel asc, step asc) rowNum
into #R_metric_queries
from GVPOperations.VID.R_Queries q with(nolock) 
inner join #R_allowedMetrics am 
	on q.step = am.metricID 
	and q.posLevel = am.posLevel 
where task like 'Metric_Insert%' 
order by posLevel



/* Grabbing the Number of Queries to Iterate Over */
declare @queryCount as int = (select count(*) from #R_metric_queries);
declare @queryNum as int = 1;

/* Vars to Be Use during each Iteration */
declare @query as nvarchar(max);
declare @poslevel as int;
declare @metricID as int;


while @queryNum <= @queryCount
begin
	/* Setting the Query, Position Level, and Metric ID to be Inserted */
		
	/* 091522 - Setting Params */
	select 
		@query = replace(r.query,'FISCALNOM',@fiscalNom) 
		,@poslevel = r.posLevel
		,@metricID = r.step
	from #R_metric_queries r with(nolock) where r.rowNum = @queryNum
	/*
	set @query = (select replace(r.query,'FISCALNOM',@fiscalNom) from #R_metric_queries r with(nolock) where r.rowNum = @queryNum);
	set @poslevel = (select r.posLevel from #R_metric_queries r with(nolock) where r.rowNum = @queryNum);
	set @metricID = (select r.step from #R_metric_queries r with(nolock) where r.rowNum = @queryNum);
	*/

	/* Trim the Metric to be Inserted */
	delete from GVPOperations.VID.R_Metrics where posLevel = @poslevel and metricID = @metricID and fiscalMth = @fiscalNom;
	/* Run Metric Insert */
	begin try
		/* Running Global Roster Section */
		/* Creating the Single Roster To Use across Sup Metrics if it doesn't exist already*/
		if(@poslevel = 2) and object_id('tempdb..##R_Sup_Roster') is null begin		
		
			declare @sup_roster_query as varchar(max);
			set @sup_roster_query = (
				select 
					replace(query,'FISCALNOM',@fiscalNom) 
				from GVPOperations.VID.R_Queries with(nolock) 
				where task = 'R_Sup_Roster') 
			execute(@sup_roster_query)
		end

		/* Running Global Roster Section */
		/* Creating the Single Roster To Use across Mgr Metrics if it doesn't exist already*/
		if(@poslevel = 3) and object_id('tempdb..##R_Mgr_Roster') is null begin		
		
			declare @mgr_roster_query as varchar(max);
			set @mgr_roster_query = (
				select 
					replace(query,'FISCALNOM',@fiscalNom) 
				from GVPOperations.VID.R_Queries with(nolock) 
				where task = 'R_Mgr_Roster') 
			execute(@mgr_roster_query)
		end

		/* Clean Up the Global when not in use */
		if(@poslevel != 3) begin
			if object_id('tempdb..##R_Mgr_Roster') is not null begin drop table ##R_Mgr_Roster end
		end
		if(@poslevel != 2) begin
			if object_id('tempdb..##R_Sup_Roster') is not null begin drop table ##R_Sup_Roster end
		end
		

		--print concat('Inserted Metric: ',cast(@metricID as varchar(2)),', for PosLevel: ',cast(@posLevel as varchar(2)))
		execute(@query);
	end try
	begin catch
		select 
			error_number() as error
			,ERROR_MESSAGE() as msg
			,@metricID as metricID
			,@posLevel as posLevel
	end catch

	set @queryNum = @queryNum + 1;
end;

set @queryCount = null;
set @queryNum = null;
set @fiscalNom = null;
set @poslevel = null;
set @metricID = null;


	IF OBJECT_ID('tempdb..#R_metric_queries') is not null begin drop table #R_metric_queries end;

-- select top(1) 1 from #R_metric_queries

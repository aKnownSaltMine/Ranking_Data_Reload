declare	@fiscalNom int = <<>>;
	
	declare @eDate date = datefromparts(left(@fiscalNom,4),right(@fiscalNom,2),28);      
	declare @sDate date = dateadd(day,1,dateadd(month,-1,@eDate));          

	/*
		Metric Rank SP

		Notes: 
			-Assuming Metric Exists becase they were video already thus no need to use Roster
			-091522: Removed Grouped MetricIDs from the Metric Needed Counts, As they are not
					required to rank overall.

	*/

	--Grabbing list of Weighted Metrics per posLevel
	--Each Emp needs the minimum Metric Count to be Ranked
	if object_id('tempdb..#R_MetricCount') is not null begin drop table #R_MetricCount end;
	if object_id('tempdb..#R_Roster') is not null begin drop table #R_Roster end;
	if object_id('tempdb..#R_groupIDs') is not null begin drop table #R_groupIDs end;

	select 
		w.groupID
	into #R_groupIDs
	from GVPOperations.VID.R_Metrics m with(nolock)
	inner join GVPOperations.VID.R_Weights w with(nolock)
		on m.metricID = w.metricID and m.fiscalMth = w.fiscalMth 
		and w.locID = 11026 and m.posLevel = w.posLevel
	where m.fiscalMth = @fiscalNom and w.metricWeight != 0.0
	and w.groupID is not null
	group by w.groupID




	select 
		m.posLevel
		,count(distinct m.metricID) metricCount
	into #R_MetricCount
	from GVPOperations.VID.R_Metrics m with(nolock)
	inner join GVPOperations.VID.R_Weights w with(nolock)
		on m.metricID = w.metricID 
		and m.fiscalMth = w.fiscalMth 
		and w.locID = 11026 
		and m.posLevel = w.posLevel
	where m.fiscalMth = @fiscalNom 
	and w.metricWeight != 0.0 
	and m.metricID not in (select groupID from #R_groupIDs g with(nolock))
	group by m.posLevel
	

	if object_id('tempdb..#metricRedistribute') is not null begin drop table #metricRedistribute end;
	
	select
		m.id
		,m.fiscalMth
		,m.posLevel
		,33 metricID
		,m.metricVal
	into #metricRedistribute
	from GVPOperations.VID.R_Metrics m with(nolock)
	inner join #R_MetricCount mc with(nolock) on m.posLevel = mc.posLevel
	left join GVPOperations.VID.R_Weights w with(nolock)
		on 
		m.metricID = 33
		and m.fiscalMth = w.fiscalMth 
		and w.locID = 11026 
		and m.posLevel = w.posLevel
	where m.posLevel = 4 and m.fiscalMth = @fiscalNom and m.metricID = 35
	group by 
		m.id
		,m.fiscalMth
		,m.posLevel
		,m.metricVal


	
	--Create the Rough Roster with needed Metric Counts
	select
		m.id
		,m.posLevel
		,mc.metricCount
		,count(distinct m.metricID) mcount
		,avg(calls.metricVal) calls
		,((avg(stfd.metricVal) /60)/60) staffedTime
		,avg(md.metricVal) metricVal
	into #R_Roster
	from GVPOperations.VID.R_Metrics m with(nolock)
	inner join #R_MetricCount mc with(nolock) on m.posLevel = mc.posLevel
	inner join GVPOperations.VID.R_Weights w with(nolock)
		on m.metricID = w.metricID and m.fiscalMth = w.fiscalMth and w.locID = 11026 and m.posLevel = w.posLevel
	/* Move calls and sftd to Roster Level */
	left join GVPOperations.VID.R_Metrics calls with(nolock) 
		on m.fiscalMth = calls.fiscalMth and m.id = calls.id and calls.metricID = 28 and m.posLevel = calls.posLevel
	left join GVPOperations.VID.R_Metrics stfd with(nolock)
		on m.fiscalMth = stfd.fiscalMth and m.id = stfd.id and stfd.metricID = 32 and stfd.posLevel = m.posLevel
	left join #metricRedistribute md with(nolock)
		on m.id = md.id and m.fiscalMth = md.fiscalMth and md.posLevel = m.posLevel

	where w.metricWeight != 0.00 and m.fiscalMth = @fiscalNom and m.metricID not in (select groupID from #R_groupIDs g with(nolock))
	group by 
		m.id
		,m.posLevel
		,mc.metricCount

	
	if object_id('tempdb..#R_MetricCount') is not null begin drop table #R_MetricCount end;

	--Clean thoes from the roster that do not have the needed Metric Count
	delete from #R_Roster where (metricVal != null and metricCount != mcount);
	delete from #R_Roster where (metricVal is null and mcount < (metricCount - 1));
	/* Metrics should Rank when within 14 days of Fiscal, After that Qualifers are used. */
	if abs(datediff(day,@sDate,getutcdate())) >= 14 begin
		delete from #R_Roster where (calls < 100 or staffedTime < 80.00);
		delete from #R_Roster where ((calls < 100 or staffedTime < 80.00) or (calls is null or staffedTime is null)) and posLevel = 4;
	end;
	
	update GVPOperations.VID.R_Metrics
	set
		metricRank = a.metricRankCalc
	from (
	select 
		m.*, mm.metricSortOrder
		,case
			/*metricSortDir is being used to determine when a dense_rank vs a regular rank should be used.*/
			/* Logic for when Rank should be used */
			when mm.metricSortDir = 0 then 		
				case
					when mm.metricSortOrder = 1 then
					/* When the Metric Rank should use regular Rank and Values are Descending */
						rank() over(
							partition by
								m.posLevel
								,m.metricID
								,m.fiscalMth
							order by
								m.metricVal desc
								)
					when mm.metricSortOrder = 0 then
					/* When the Metric Rank should use regular Rank and Values are Ascending */
						rank() over(
						partition by
							m.posLevel
							,m.metricID
							,m.fiscalMth
						order by
							m.metricVal asc
							)
					else null 
				end 
			/* Logic for when Dense Rank should be used */
			when mm.metricSortDir = 1 then
				case
					when mm.metricSortOrder = 1 then 
					/* When the Metric Rank should use DENSE Rank and Values are Descending */
						dense_rank() over(
							partition by
								m.posLevel
								,m.metricID
								,m.fiscalMth
							order by
								m.metricVal desc
								)
					when mm.metricSortOrder = 0 then
					/* When the Metric Rank should use DENSE Rank and Values are Ascending */
						dense_rank() over(
						partition by
							m.posLevel
							,m.metricID
							,m.fiscalMth
						order by
							m.metricVal asc
							)
					else null 
				end
			else null 
		end
		metricRankCalc


	from GVPOperations.VID.R_Metrics m with(nolock)
	--Inner to grab only thoes with the correct count of Weighted metrics
	inner join #R_Roster r with(nolock) on m.id = r.id and m.posLevel = r.posLevel
	inner join GVPOperations.VID.R_Metric_Map mm with(nolock)
		on m.metricID = mm.id
	inner join GVPOperations.VID.R_Weights w with(nolock)
		on m.metricID = w.metricID and m.fiscalMth = w.fiscalMth and m.posLevel = w.posLevel and w.locID = 11026
	where m.fiscalMth = @fiscalNom and w.metricWeight != 0.00
	) as a
	where GVPOperations.VID.R_Metrics.id = a.id and 
		GVPOperations.VID.R_Metrics.fiscalMth = a.fiscalMth and 
		GVPOperations.VID.R_Metrics.metricID = a.metricID and
		GVPOperations.VID.R_Metrics.posLevel = a.posLevel and 
		case
			when a.metricID = 18 and a.posLevel = 2 then 0
			when a.metricID = 42 and a.posLevel = 6 then 0
			else 1
		end = 1
	
	
	if object_id('tempdb..#R_groupIDs') is not null begin drop table #R_groupIDs end;
	if object_id('tempdb..#R_Roster') is not null begin drop table #R_Roster end;

-- select Top(1) 1 from GVPOperations.VID.R_Ranking;



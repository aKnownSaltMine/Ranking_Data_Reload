declare	@fiscalNom int = <<>>;

	declare @eDate date = datefromparts(left(@fiscalNom,4),right(@fiscalNom,2),28);      
	declare @sDate date = dateadd(day,1,dateadd(month,-1,@eDate)); 



	if object_id('tempdb..#R_Ranks') is not null begin drop table #R_Ranks end;

	--Grabbing list of Weighted Metrics per posLevel
	--Each Emp needs the minimum Metric Count to be Ranked
	if object_id('tempdb..#R_MetricCount') is not null begin drop table #R_MetricCount end;
	if object_id('tempdb..#R_Roster') is not null begin drop table #R_Roster end;
	if object_id('tempdb..#R_groupIDs') is not null begin drop table #R_groupIDs end;

	select 
		w.groupID
		,w.fiscalMth
		,w.posLevel
	into #R_groupIDs
	from GVPOperations.VID.R_Metrics m with(nolock)
	inner join GVPOperations.VID.R_Weights w with(nolock)
		on m.metricID = w.metricID and m.fiscalMth = w.fiscalMth 
		and w.locID = 11026 and m.posLevel = w.posLevel
	where m.fiscalMth = @fiscalNom and w.metricWeight != 0.0
	and w.groupID is not null
	group by w.groupID, w.fiscalMth, w.posLevel


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
		m.posLevel
		,m.fiscalMth
		,count(distinct m.metricID) metricCount
	into #R_MetricCount
	from GVPOperations.VID.R_Metrics m with(nolock)
	inner join GVPOperations.VID.R_Weights w with(nolock)
		on m.metricID = w.metricID 
		and m.fiscalMth = w.fiscalMth 
		and w.locID = 11026 
		and m.posLevel = w.posLevel
	left join #R_groupIDs g with(nolock) 
		on m.metricID = g.groupID 
		and m.fiscalMth = g.fiscalMth 
		and m.posLevel = g.posLevel
	where m.fiscalMth = @fiscalNom 
	and w.metricWeight != 0.0 
	and g.groupID is null
	group by m.posLevel, m.fiscalMth
	



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

	delete from #metricRedistribute where metricVal is null or metricVal > 1


	--Create the Rough Roster with needed Metric Counts
	select
		m.id
		,m.posLevel
		,mc.metricCount
		,count(distinct m.metricID) mcount
		,avg(calls.metricVal) calls
		,((avg(stfd.metricVal) /60)/60) staffedTime
		,case 
			when avg(md.metricVal) is null then 0
			when avg(md.metricVal) is not null then 0.025 
		end metricRedsitribute
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
	delete from #R_Roster where (metricRedsitribute != null and metricCount != mcount);
	delete from #R_Roster where (metricRedsitribute is null and mcount < (metricCount - 1));
	/* Metrics should Rank when within 14 days of Fiscal, After that Qualifers are used. */
	if abs(datediff(day,@sDate,getutcdate())) >= 14 begin

		/* Removing thoes that dont meet Minimums */
		
		delete from #R_Roster
			where 
				case 
					/* Leads have moved to using 250 Calls as a Minimum as of Fiscal April '23 */
					when #R_Roster.posLevel = 4 and @fiscalNom >= 202305 then 
						case
							when ((#R_Roster.calls < 250) or (#R_Roster.staffedTime < 80.0)) or ((#R_Roster.calls = null) or (#R_Roster.staffedTime = null)) then 1
							else 0
						end
					else
						case
							when ((#R_Roster.calls < 100) or (#R_Roster.staffedTime < 80.0)) or ((#R_Roster.calls = null) or (#R_Roster.staffedTime = null)) then 1
							else 0
						end
				end = 1;
		
	end;

	delete from GVPOperations.VID.R_Ranking where fiscalMth = @fiscalNom

	;;with wTotals as (
	select 
		m.*
		,r.metricRedsitribute
		,case when r.metricRedsitribute is not null then
			case 
				when m.metricID = 33 then null
				else
					cast(m.metricRank as float) * (w.metricWeight + case when r.metricRedsitribute is null then 0 else r.metricRedsitribute end) 
				end
			when r.metricRedsitribute is null then 
				cast(m.metricRank as float) * (w.metricWeight + case when r.metricRedsitribute is null then 0 else r.metricRedsitribute end) 
			end wTotal
	from GVPOperations.VID.R_Metrics m with(nolock)
	inner join #R_Roster r with(nolock) on m.id = r.id and m.posLevel = r.posLevel
	inner join GVPOperations.VID.R_Weights w with(nolock)
		on m.metricID = w.metricID 
		and m.fiscalMth = w.fiscalMth 
		and m.posLevel = w.posLevel 
		and w.metricWeight != 0.00
		and w.locID = 11026
	where m.fiscalMth = @fiscalNom
	/* 030823 - Group Weighted totals were being included into the overall. Removed from Calc */
	and m.metricRank is not null and m.metricID not in (select groupID from #R_groupIDs g with(nolock) group by groupID))

	select 
		w.id
		,w.fiscalMth
		,w.posLevel
		,sum(w.wTotal) wTotal
	into #R_Ranks
	from wTotals w with(nolock)
	group by 
		w.id
		,w.fiscalMth
		,w.posLevel




	-- Update Existing Weighted Total Values
	update r 
		set r.weightTotal = w.wTotal
	from GVPOperations.VID.R_Ranking r with(nolock)
	inner join #R_Ranks w 
		on r.id = w.id 
		and r.fiscalMth = w.fiscalMth 
		and r.posLevel = w.posLevel



	-- Insert new Weighted Total Values
	insert GVPOperations.VID.R_Ranking(id, fiscalMth, posLevel, weightTotal)
	select 
		id
		,fiscalMth
		,posLevel
		,wTotal weightTotal
	from #R_Ranks r with(nolock)
	where concat(r.id,r.fiscalMth,r.posLevel) not in (
	select concat(a.id,a.fiscalMth,a.posLevel) from GVPOperations.VID.R_Ranking a with(nolock)
	where a.fiscalMth = @fiscalNom)
	and r.fiscalMth = @fiscalNom

	
	update r
		set r.fiscalPctRank = w.fiscalPctRank, r.fiscalRank = w.fiscalRank
	from GVPOperations.VID.R_Ranking r with(nolock)
	inner join (
	select r.*,
	rank()
		over(
			partition by 
				r.posLevel
				,r.fiscalMth
			order by
				r.weightTotal asc) rnk	
	,PERCENT_RANK()
		over(
			partition by 
				r.posLevel
				,r.fiscalMth
			order by
				r.weightTotal desc) rnkpct
	from GVPOperations.VID.R_Ranking r with(nolock)) w
	on r.id = w.id 
	and r.fiscalMth = w.fiscalMth 
	and r.posLevel = w.posLevel


	update GVPOperations.VID.R_Ranking
	set fiscalRank = r.rnk, fiscalPctRank = r.rnkpct
	from (
	select r.*,
	rank()
		over(
			partition by 
				r.posLevel
				,r.fiscalMth
			order by
				r.weightTotal asc) rnk	
	,PERCENT_RANK()
		over(
			partition by 
				r.posLevel
				,r.fiscalMth
			order by
				r.weightTotal desc) rnkpct
	from GVPOperations.VID.R_Ranking r with(nolock)
	) as r
	where GVPOperations.VID.R_Ranking.id = r.id
	and GVPOperations.VID.R_Ranking.fiscalMth = r.fiscalMth
	and GVPOperations.VID.R_Ranking.posLevel = r.posLevel

	
	if object_id('tempdb..#R_Ranks') is not null begin drop table #R_Ranks end;
	if object_id('tempdb..#R_Roster') is not null begin drop table #R_Roster end;

-- select Top(1) 1 from GVPOperations.VID.R_Ranking;
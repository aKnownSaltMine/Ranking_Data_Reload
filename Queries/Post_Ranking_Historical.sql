declare @runDateTime as datetime = dateadd(hour,-5,getutcdate());
declare	@fiscalNom int = <<>>;
/* Historical Save */
insert into GVPOperations.VID.R_Metrics_Hist(id, fiscalMth, metricID, metricVal, metricRank, rankType, posLevel, ChangedOn)
select
	id
	,fiscalMth
	,metricID
	,metricVal
	,metricRank
	,rankType
	,posLevel
	,@runDateTime
from GVPOperations.VID.R_Metrics m with(nolock)
where m.fiscalMth = @fiscalNom

insert into GVPOperations.VID.R_Ranking_Hist(id, fiscalMth, weightTotal, fiscalRank, fiscalPctRank, rankType, posLevel, ChangedOn)
select
	id
	,fiscalMth
	,weightTotal
	,fiscalRank
	,fiscalPctRank
	,rankType
	,posLevel
	,@runDateTime
from GVPOperations.VID.R_Ranking r with(nolock)
where r.fiscalMth = @fiscalNom

/* Updated to Use Var as datetime stamp rather than Job Token */
update a
set a.Status_TS = @runDateTime
from GVPOperations.VID.JOB_StatusTracking a with(nolock)
where a.JobToken = 'Ranking.Daily'

-- select Top(1) 1 from GVPOperations.VID.R_Ranking;
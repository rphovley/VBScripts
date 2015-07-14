select cpay.business_center_id as rep_id, concat(d.first_name,d.last_name) as rep_name, cpay.value as commission_payout, cpd.job_id, cr.name as commission_type

from commission_payouts as cpay, commission_requirements as cr, commission_periods as cper, distributors as d,
commission_payout_details as cpd
	join job_data jd on jd.job_id=cpd.job_id
	join office_reps on office_reps.distributor_id = cpd.business_center_id
	join offices as o on o.office_id = office_reps.office_id
	join office_managers as om on om.office_id = o.office_id
	join office_regionals as ore on ore.office_id = o.office_id
	join office_dvps as od on od.office_id = o.office_id
	join office_recruiters as orc on orc.distributor_id = cpd.business_center_id
		where cpay.commission_period_id=1 and cpay.commission_period_id=cper.commission_period_id
			and cpay.commission_requirements_id=cr.commission_requirements_id
			and cpay.commission_payout_id = cpd.commission_payout_id
			and cpay.business_center_id=d.distributor_id
			order by rep_id, cr.name

select 
cpay.business_center_id as rep_id, concat(d.first_name,d.last_name) as rep_name, cpay.value as commission_payout, cpd.job_id, cr.name as commission_type
	from commission_payouts as cpay, commission_requirements as cr, commission_periods as cper, distributors as d,
	commission_payout_details as cpd
		left outer join job_data jd on jd.job_id=cpd.job_id
		where cpay.commission_period_id=1 and cpay.commission_period_id=cper.commission_period_id
		and cpay.commission_requirements_id=cr.commission_requirements_id
		and cpay.commission_payout_id = cpd.commission_payout_id
		and cpay.business_center_id=d.distributor_id
			order by rep_id, cr.name


select
	cpay.commission_payout_id as 'payout_id', cr.name as 'commission_type', concat(d1.first_name,d1.last_name) as 'upline_rep', concat(d2.first_name, d2.last_name)as 'downline_rep',
	cpay.value as 'payment_amount', j.system_size,  j.job_id, j.region, o.name as 'office_name'
from commission_payouts as cpay
	left join commission_payout_details cpd on cpay.commission_payout_id       = cpd.commission_payout_id
	left join commission_requirements cr    on cpay.commission_requirements_id = cr.commission_requirements_id
	left join commission_periods cper       on cpay.commission_period_id       = cper.commission_period_id
	left join distributors d1               on cpay.business_center_id         = d1.distributor_id
	left join job_data j                    on j.job_id                        = cpd.job_id
	left join distributors d2               on j.distributor_id                = d2.distributor_id
	inner join office_reps orep             on j.distributor_id                = orep.distributor_id
	left join offices o                     on o.office_id                     = orep.office_id
where cpay.commission_period_id = 7
   

select
	cpay.commission_payout_id as 'payout_id', concat(dpay.first_name,dpay.last_name) as 'upline_rep', cpay.business_center_id as 'upline_rep_id', cr.name as 'commission_type', cpay.value as 'payment_amount',
	j.system_size , cpd.job_id, concat(drep.first_name,drep.last_name) as 'upline_rep', j.distributor_id as 'downline_rep_id'
from commission_payouts as cpay
	inner join commission_payout_details cpd on cpay.commission_payout_id       = cpd.commission_payout_id
	inner join distributors as dpay          on cpay.business_center_id         = dpay.distributor_id
	inner join job_data j                    on j.job_id                        = cpd.job_id
	inner join commission_requirements cr    on cpay.commission_requirements_id = cr.commission_requirements_id
	left join office_reps orep               on j.distributor_id                = orep.distributor_id
	inner join distributors as drep          on j.distributor_id                = drep.distributor_id
where cpay.commission_period_id = 7


select
	cpay.commission_payout_id as 'payout_id', concat(dpay.first_name,dpay.last_name) as 'upline_rep', cpay.business_center_id as 'upline_rep_id', concat(drep.first_name,drep.last_name) as 'downline_rep', j.distributor_id as 'downline_rep_id',
	j.customer_name, j.job_id, j.date_sold, " " as 'date_paid', " " as 'reason', " " as 'status', " " as 'sub_status',
	cr.name as 'commission_type', CAST((CAST(cpay.value as decimal(6,2))/CAST(j.system_size as decimal(4,2))) as decimal(4,2)) as "rate", j.system_size, cpay.value as 'payment_amount'
	
from commission_payouts as cpay
	inner join commission_payout_details cpd on cpay.commission_payout_id       = cpd.commission_payout_id
	inner join distributors as dpay          on cpay.business_center_id         = dpay.distributor_id
	inner join job_data j                    on j.job_id                        = cpd.job_id
	inner join commission_requirements cr    on cpay.commission_requirements_id = cr.commission_requirements_id
	left join office_reps orep               on j.distributor_id                = orep.distributor_id
	inner join distributors as drep          on j.distributor_id                = drep.distributor_id
where cpay.commission_requirements_id = 3 AND j.job_id = '2029067' OR j.job_id = '2032731' OR j.job_id = '2025199' OR j.job_id = '2026660' and cpay.commission_payout_type_id = 15

select
	cpay.commission_payout_id as 'payout_id', concat(dpay.first_name,dpay.last_name) as 'upline_rep', cpay.business_center_id as 'upline_rep_id', concat(drep.first_name,drep.last_name) as 'downline_rep', j.distributor_id as 'downline_rep_id',
	j.customer_name, j.job_id, j.date_sold, " " as 'date_paid', " " as 'reason', " " as 'status', " " as 'sub_status',
	cr.name as 'commission_type', CAST((CAST(cpay.value as decimal(6,2))/CAST(j.system_size as decimal(4,2))) as decimal(4,2)) as "rate", j.system_size, cpay.value as 'payment_amount'
	
from commission_payouts as cpay
	inner join commission_payout_details cpd on cpay.commission_payout_id       = cpd.commission_payout_id
	inner join distributors as dpay          on cpay.business_center_id         = dpay.distributor_id
	inner join job_data j                    on j.job_id                        = cpd.job_id
	inner join commission_requirements cr    on cpay.commission_requirements_id = cr.commission_requirements_id
	left join office_reps orep               on j.distributor_id                = orep.distributor_id
	inner join distributors as drep          on j.distributor_id                = drep.distributor_id
where cpay.commission_requirements_id = 3 AND cpay.commission_payout_id = 8408 OR cpay.commission_payout_id = 8407 OR cpay.commission_payout_id = 8353 OR cpay.commission_payout_id = 8352 and cpay.commission_payout_type_id = 15

commission_payouts cpay
commission_requirements cr
commission_payout_details cpd
commission_periods cper
distributors d
job_data j
offices o
office_managers om
office_regionals oreg
office_dvps od
office_recruiters orec
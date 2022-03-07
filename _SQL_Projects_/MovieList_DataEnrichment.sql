
-- Company name starting with 'The' -------------------------------------------------------------------------------

select*
from portfolio_project.dbo.Data_movieList
where company like'%the %'



-- Company producing more than 19 movies --------------------------------------------------------------------------

select company, count(name) as amount
from portfolio_project.dbo.Data_movieList
group by company
having count(name) > 19
order by amount desc



--  Company producing more than 1 but  less than 6 movies --------------------------------------------------------

select company, count(name) as amount
from portfolio_project.dbo.Data_movieList
group by company
having count(name) >= 2 and count(name) <= 5
order by amount



--  Only the company Names meeting above criteria -----------------------------------------------------------------

select a.company
from
	(select company, count(name) as amount
	from portfolio_project.dbo.Data_movieList
	group by company
	having count(name) > 19) a
order by amount desc


--  Select All data from above mentioned companies -----------------------------------------------------------------

select*
from portfolio_project.dbo.Data_movieList
where company in 
				(select sub.company from
										(select company, count(name) as amount
										from portfolio_project.dbo.Data_movieList
										group by company
										having count(name) >= 2 and count(name) <= 5) 
										sub)
and budget is not null 
and company is not null
order by score desc


-----------------------------------------------------------------------------------------------------------------------------------------------------------------
-- Copy All data to a Temporary table 1 and Updating Company Names ----------------------------------------------------------------------------------------------
-----------------------------------------------------------------------------------------------------------------------------------------------------------------


drop table if exists #temp_movieList1
create table #temp_movieList1 (
							name nvarchar(255), 
							rating nvarchar(255),
							genre nvarchar(255),
							year float,
							released nvarchar(255),
							score float,
							votes float,
							director nvarchar(255),
							writer nvarchar(255),
							star nvarchar(255),
							country nvarchar(255),
							budget float,
							gross float,
							company nvarchar(255),
							runtime float
							)

insert into #temp_movieList1
select* from portfolio_project.dbo.Data_movieList



select a.company
from
	(select company, count(name) as amount
	from #temp_movieList1
	group by company
	having count(name) > 5) a



select* from #temp_movieList1



update #temp_movieList1
set company = 'Dreamworks'
where company like '%dreamworks%'

update #temp_movieList1
set company = 'Twentieth Century'
where company like '%twentieth%'

update #temp_movieList1
set company = 'Fox Pictures'
where company like '%fox%'

update #temp_movieList1
set company = 'Paramount Pictures'
where company like '%paramount%'

update #temp_movieList1
set company = 'Lions Gate'
where company like '%lion%'

update #temp_movieList1
set company = 'Walt Disney'
where company like '%walt%'

update #temp_movieList1
set company = 'Warner Bros'
where company like '%warner%'

update #temp_movieList1
set company = 'Alliance Communications'
where company like '%alliance%'

update #temp_movieList1
set company = 'BBC Films'
where company like '%BBC%'

update #temp_movieList1
set company = 'CBS Films'
where company like '%CBS%'

update #temp_movieList1
set company = 'Imagine Entertainment'
where company like '%imagine%'

update #temp_movieList1
set company = 'American Zoetrope'
where company like '%zoetrope%'


-----------------------------------------------------------------------------------------------------------------------------------------------------------------
-- Copy All data to a Temporary table 2 and Extracting Month, Year & Country of Release and Adding counting in Millions columns ---------------------------------
-----------------------------------------------------------------------------------------------------------------------------------------------------------------


drop table if exists #temp_movieList2
create table #temp_movieList2 (
							name nvarchar(255), 
							rating nvarchar(255),
							genre nvarchar(255),
							year float,
							released nvarchar(255),
							score float,
							votes float,
							director nvarchar(255),
							writer nvarchar(255),
							star nvarchar(255),
							country nvarchar(255),
							budget float,
							gross float,
							company nvarchar(255),
							runtime float
							)

insert into #temp_movieList2
select* from #temp_movieList1



alter table #temp_movieList2
add month_released nvarchar(255);

update #temp_movieList2
set month_released = substring(released, 1, 3)



alter table #temp_movieList2
add year_released nvarchar(255);

update #temp_movieList2
set year_released = substring( parsename( replace(released, ',', '.') , 1),     2,    5)


--select convert (int, #temp_movieList1.year_released) from #temp_movieList1
--cast(year_released as float)


alter table #temp_movieList2
add country_released nvarchar(255);

update #temp_movieList2
set country_released = substring( parsename( replace(released, '(', '.') , 1),     1,    len(parsename( replace(released, '(', '.') , 1))-1)



------------------------------------------------------------------------------------------------------------------


alter table #temp_movieList2
add budget_million float;

update #temp_movieList2
set budget_million = budget/1000000



alter table #temp_movieList2
add gross_million float;

update #temp_movieList2
set gross_million = gross/1000000



alter table #temp_movieList2
add votes_million float;

update #temp_movieList2
set votes_million = votes/1000000



alter table #temp_movieList2
add points_million float;

update #temp_movieList2
set points_million = votes_million*score



select * from #temp_movieList2


---------------------------------------------------------------------------------------------------------------------

/*
select *
From #temp_movieList2
where rating is null
or votes is null
or star is null


delete
from #temp_movieList2
where released is null
or company is null
or budget is null
or gross is null
or runtime is null
*/
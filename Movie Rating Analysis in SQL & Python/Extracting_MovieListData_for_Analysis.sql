-----------------------------------------------------------------------------------------------------------------------------------------------------------------
/* Filtering and Extracting Data of the Movie Companies producing more than 5 movies */
-----------------------------------------------------------------------------------------------------------------------------------------------------------------

select* from portfolio_project.dbo.movieList_updated
order by votes desc


drop table if exists #temp_movieList3
create table #temp_movieList3 (
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
							runtime float,
							month_released nvarchar(255),
							year_released nvarchar(255),
							country_released nvarchar(255),
							budget_million float,
							gross_million float,
							votes_million float,
							points_million float
							)

insert into #temp_movieList3
select*
from portfolio_project.dbo.movieList_updated
where company in 
				(select sub.company from
										(select company, count(name) as amount
										from portfolio_project.dbo.movieList_updated
										group by company
										having count(name) > 5  --order by amount desc
										) 
				sub)



select* from #temp_movieList3
order by votes desc



select*
from #temp_movieList3
where rating is null
or year_released like '%(%'
or month_released like '%19%'


delete from #temp_movieList3 
where month_released like '%19%'

update #temp_movieList3
set rating = 'None'
where rating like '%rated%'
or rating is null


select company, name, count (name) over (partition by company) as company_count 
from #temp_movieList3
order by company_count desc



-- company, genre, country, director, star, rating ------------------------------------------------------------------

drop table if exists #temp_company
create table #temp_company (
							company nvarchar(255), 
							company_count int
							)
select* from #temp_company order by company_count desc

insert into #temp_company
select company, count(name)
from #temp_movieList3
group by company


drop table if exists #temp_companys
create table #temp_companys (
							company nvarchar(255),
							company_count int,
							company_rank int
							)
select* from #temp_companys order by company_rank

insert into #temp_companys
select company, company_count, rank() over(order by company_count desc) 
from #temp_company

--------------------------------------------------------------------------------------------------------------------

drop table if exists #temp_genre
create table #temp_genre (
							genre nvarchar(255), 
							genre_count int
							)
select* from #temp_genre order by genre_count desc

insert into #temp_genre
select genre, count(name)
from #temp_movieList3
group by genre


drop table if exists #temp_genres
create table #temp_genres (
							genre nvarchar(255),
							genre_count int,
							genre_rank int
							)
select* from #temp_genres order by genre_rank

insert into #temp_genres
select genre, genre_count, rank() over(order by genre_count desc) 
from #temp_genre

--------------------------------------------------------------------------------------------------------------------

drop table if exists #temp_country
create table #temp_country (
							country nvarchar(255), 
							country_count int
							)
select* from #temp_country order by country_count desc

insert into #temp_country
select country, count(name)
from #temp_movieList3
group by country


drop table if exists #temp_countrys
create table #temp_countrys (
							country nvarchar(255),
							country_count int,
							country_rank int
							)
select* from #temp_countrys order by country_rank

insert into #temp_countrys
select country, country_count, rank() over(order by country_count desc) 
from #temp_country

--------------------------------------------------------------------------------------------------------------------

drop table if exists #temp_director
create table #temp_director (
							director nvarchar(255), 
							director_count int
							)
select* from #temp_director order by director_count desc

insert into #temp_director
select director, count(name)
from #temp_movieList3
group by director


drop table if exists #temp_directors
create table #temp_directors (
							director nvarchar(255),
							director_count int,
							director_rank int
							)
select* from #temp_directors order by director_rank

insert into #temp_directors
select director, director_count, rank() over(order by director_count desc) 
from #temp_director

--------------------------------------------------------------------------------------------------------------------

drop table if exists #temp_star
create table #temp_star (
							star nvarchar(255), 
							star_count int
							)
select* from #temp_star order by star_count desc

insert into #temp_star
select star, count(name)
from #temp_movieList3
group by star


drop table if exists #temp_stars
create table #temp_stars (
							star nvarchar(255),
							star_count int,
							star_rank int
							)
select* from #temp_stars order by star_rank

insert into #temp_stars
select star, star_count, rank() over(order by star_count desc) 
from #temp_star

--------------------------------------------------------------------------------------------------------------------

drop table if exists #temp_rating
create table #temp_rating (
							rating nvarchar(255), 
							rating_count int
							)
select* from #temp_rating order by rating_count desc

insert into #temp_rating
select rating, count(name)
from #temp_movieList3
group by rating


drop table if exists #temp_ratings
create table #temp_ratings (
							rating nvarchar(255),
							rating_count int,
							rating_rank int
							)
select* from #temp_ratings order by rating_rank

insert into #temp_ratings
select rating, rating_count, rank() over(order by rating_count desc) 
from #temp_rating


--------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------


select *
	from #temp_movieList3	a
join #temp_companys	b
	on a.company =			b.company
join #temp_genres c
	on a.genre =			c.genre
join #temp_directors d
	on a.director =			d.director
join #temp_stars e
	on a.star =				e.star
join #temp_countrys f
	on a.country =			f.country
join #temp_ratings g
	on a.rating =			g.rating
/*
where a.rating is null
or year_released like '%(%' 
or month_released like '%19%'
*/
where a.budget is not null 
and a.gross is not null
order by year

---------------------------------------------------------------------------------------------------------------------

drop table if exists #movieList_refined

Select 
	name,
	runtime,
	year,
	released,
	month_released,
	year_released,
	country_released,
	budget_million,
	gross_million,
	votes_million,
	points_million,
	score,
	writer,
	a.director,
	director_count,
	director_rank,
	a.star,
	star_count,
	star_rank,
	a.country,
	country_count,
	country_rank,
	a.company,
	company_count,
	company_rank,
	a.genre,
	genre_count,
	genre_rank,
	a.rating,
	rating_count,
	rating_rank
into 
	#movieList_refined
from 
	#temp_movieList3	a
	join #temp_companys	b
		on a.company =		b.company
	join #temp_genres c
		on a.genre =		c.genre
	join #temp_directors d
		on a.director =		d.director
	join #temp_stars e
		on a.star =			e.star
	join #temp_countrys f
		on a.country =		f.country
	join #temp_ratings g
		on a.rating =		g.rating
/*
where a.rating is null
or year_released like '%(%' 
or month_released like '%19%'
*/
where a.budget is not null 
and a.gross is not null
and a.runtime is not null
order by year


Select * From #movieList_refined


--------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------


select year, score, votes_million, budget_million, gross_million
from #temp_movieList3

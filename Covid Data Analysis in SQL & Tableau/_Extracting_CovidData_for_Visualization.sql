
----------------------------------------------------------------------------------------------------------------------------------------------------------
--------------------------------------------------- Statistics on Covid Infection until February, 2022 ---------------------------------------------------
----------------------------------------------------------------------------------------------------------------------------------------------------------



-- 1. Total Cases, Deaths and Percentage

Select SUM(new_cases) as cases_total_daily, SUM(cast(new_deaths as int)) as death_total_daily, SUM(cast(new_deaths as int))/SUM(new_cases)*100 as death_percent_daily
From portfolio_project.dbo.covid_infect
Where continent is not null 
Order by cases_total_daily, death_total_daily



-- 2. Daily Death counts in every Continent

Select location, SUM(cast(new_deaths as int)) as death_total_daily
From portfolio_project.dbo.covid_infect
Where continent is null 
And location not in ('World', 'European Union', 'International')
And location not like '%income%'
Group by location
Order by death_total_daily desc



-- 3. Infection rates and Population of each Country

Select location, population, MAX(total_cases) as infected_count_max,  Max((total_cases/population))*100 as infected_percent_max
From portfolio_project.dbo.covid_infect
Group by location, population
Order by infected_percent_max desc



-- 4. Time-Series of Daily Cases in each Country

Select location, population, date, MAX(total_cases) as infected_count_max,  Max((total_cases/population))*100 as infected_percent_max
From portfolio_project.dbo.covid_infect
Group by location, population, date
Order by infected_percent_max desc



----------------------------------------------------------------------------------------------------------------------------------------------------------
-------------------------------------------------- Statistics on Covid Vaccination until February, 2022 --------------------------------------------------
----------------------------------------------------------------------------------------------------------------------------------------------------------



-- 5. Time-series of Daily Vaccinations in each Country

Select inf.continent, inf.location, inf.date, inf.population, 
	MAX(vac.total_vaccinations) as RollingPeopleVaccinated --, (RollingPeopleVaccinated/population)*100
From portfolio_project.dbo.covid_infect inf
Join portfolio_project.dbo.covid_vaccine vac
	On inf.location = vac.location
	And inf.date = vac.date
Where inf.continent is not null 
Group by inf.continent, inf.location, inf.date, inf.population
Order by inf.date desc--, RollingPeopleVaccinated desc



-- 6. Total Population, Vaccination and Percentage 

With pop_vs_vac (Continent, Location, Date, Population, New_Vaccinations, RollingPeopleVaccinated) 
as	(	
	Select	inf.continent, inf.location, inf.date, inf.population, 
			vac.new_vaccinations, 
			SUM(CONVERT(bigint,vac.new_vaccinations)) OVER (Partition by inf.Location Order by inf.location, inf.Date) as RollingPeopleVaccinated 
			--, (RollingPeopleVaccinated/population)*100
	From portfolio_project.dbo.covid_infect inf
	Join portfolio_project.dbo.covid_vaccine vac
		On inf.location = vac.location
		And inf.date = vac.date
	Where inf.continent is not null
	)
Select *, (RollingPeopleVaccinated/Population)*100 as PercentPeopleVaccinated
From pop_vs_vac



-- 7. Canadian Stats on Covid

Select inf.date, population, total_cases as accumulated_case, total_deaths as accumulated_death, 
			SUM(CONVERT(bigint,vac.new_vaccinations)) OVER (Partition by inf.Location Order by inf.location, inf.Date) as accumulated_vaccination
From portfolio_project.dbo.covid_infect inf
Join portfolio_project.dbo.covid_vaccine vac
	On inf.location = vac.location
	And inf.date = vac.date
Where inf.location like '%canada%'
Order by inf.date



-- 8. Accumulated Cases and Deaths of all Countries after every Month (Data Extarction for PowerBI)

Select location as Country, cast(date as varchar), total_cases as Accumulated_Case, total_deaths as Accumulated_Death, population as Populations
From portfolio_project.dbo.covid_infect
Where cast(date as varchar) like '%Feb  1 2022%'
-- Where cast(date as varchar) like '%  1 20%'
and continent is not null 
Order by location, date desc
/*
Covid 19 Data Exploration 
Skills used: Aggregate Functions, Windows Functions, Converting Data Types, Joins, CTE, Creating Views, Temp Tables
*/  

-- Checking the Data

Select *
From portfolio_project.dbo.covid_infect
Order by location, date


-- Selecting Data Columns that are useful for the Project

Select location, date, total_cases, new_cases, total_deaths, population
From portfolio_project.dbo.covid_infect
Where continent is not null 
Order by 1,2 desc


-- Total Cases vs Total Deaths
-- Shows likelihood of dying if you contract covid in your country

Select location, date, total_cases,total_deaths, (total_deaths/total_cases)*100 as death_percent
From portfolio_project.dbo.covid_infect
Where location like '%canada%'
and continent is not null 
Order by 1,2 desc


-- Total Cases vs population
-- Shows what percentage of population infected with Covid

Select location, date, population, total_cases,  (total_cases/population)*100 as infected_percent
From portfolio_project.dbo.covid_infect
Where location like '%canada%'
and continent is not null 
Order by 1,2 desc


-- Countries with Highest Infection Rate compared to population

Select location, population, MAX(total_cases) as infected_count_max,  Max((total_cases/population))*100 as infected_percent_max
From portfolio_project.dbo.covid_infect
Where population > 5000000
Group by location, population
Order by infected_percent_max desc


-- Countries with Highest Death Count per population

Select location, MAX(cast(Total_deaths as int)) as death_count_max
From portfolio_project.dbo.covid_infect
Where continent is not null 
Group by location
Order by death_count_max desc




-- BREAKING THINGS DOWN BY CONTINENT

-- Showing contintents with the highest death count per population

-- Method #1
Select continent, MAX(cast(Total_deaths as int)) as death_count_max
From portfolio_project.dbo.covid_infect
Where continent is not null 
Group by continent
Order by death_count_max desc

-- Method #2
Select location, MAX(cast(Total_deaths as int)) as death_count_max
From portfolio_project.dbo.covid_infect
Where continent is null 
and location not like '%income%'
Group by location
Order by death_count_max desc



-- GLOBAL NUMBERS

Select SUM(new_cases) as total_cases_daily, SUM(cast(new_deaths as int)) as death_total_daily, SUM(cast(new_deaths as int))/SUM(new_cases)*100 as death_percent_daily
From portfolio_project.dbo.covid_infect
Where continent is not null 


Select date, SUM(new_cases) as total_cases_daily, SUM(cast(new_deaths as int)) as death_total_daily, SUM(cast(new_deaths as int))/SUM(new_cases)*100 as death_percent_daily
From portfolio_project.dbo.covid_infect
Where location like '%canada%'
and continent is not null 
Group By date
Order by date desc



-- Total population vs Vaccinations
-- Shows Percentage of population that has recieved at least one Covid Vaccine

Select	inf.continent, inf.location, inf.date, inf.population, 
		vac.new_vaccinations
From portfolio_project.dbo.covid_infect inf
Join portfolio_project.dbo.covid_vaccine vac
	On inf.location = vac.location
	and inf.date = vac.date
Where inf.continent is not null 
Order by 1,2,3



Select	inf.continent, inf.location, inf.date, inf.population, 
		vac.new_vaccinations, 
		SUM(CONVERT(bigint,vac.new_vaccinations)) OVER (Partition by inf.location Order by inf.location, inf.date) as vaccinated_total
From portfolio_project.dbo.covid_infect inf
Join portfolio_project.dbo.covid_vaccine vac
	On inf.location = vac.location
	and inf.date = vac.date
Where inf.continent is not null 
Order by 2,3 desc


-- Using CTE to perform Calculation on Partition By in previous query

With PopvsVac (continent, location, date, population, new_vaccinations, vaccinated_total) as (

	Select	inf.continent, inf.location, inf.date, inf.population, 
			vac.new_vaccinations, 
			SUM(CONVERT(bigint,vac.new_vaccinations)) OVER (Partition by inf.location Order by inf.location, inf.date) as vaccinated_total
	From portfolio_project.dbo.covid_infect inf
	Join portfolio_project.dbo.covid_vaccine vac
		On inf.location = vac.location
		and inf.date = vac.date
	Where inf.continent is not null)

Select *, (vaccinated_total/population)*100 as vaccinated_percent
From PopvsVac



-- Using Temp Table to perform Calculation on Partition By in previous query

DROP Table if exists #Percent_Population_Vaccinated

Create Table #Percent_Population_Vaccinated (
	continent nvarchar(255),
	location nvarchar(255),
	date datetime,
	population numeric,
	new_vaccinations numeric,
	vaccinated_total numeric)



Insert into #Percent_Population_Vaccinated

Select	inf.continent, inf.location, inf.date, inf.population, 
		vac.new_vaccinations, 
		SUM(CONVERT(bigint,vac.new_vaccinations)) OVER (Partition by inf.location Order by inf.location, inf.date) as vaccinated_total
From portfolio_project.dbo.covid_infect inf
Join portfolio_project.dbo.covid_vaccine vac
	On inf.location = vac.location
	and inf.date = vac.date
Where inf.continent is not null 
--Order by 2,3



Select *, (vaccinated_total/population)*100 as vaccinated_percent
From #Percent_Population_Vaccinated




-- Creating View to store data for later visualizations

CREATE VIEW PercentPopulationVaccination as
--Select *
--From portfolio_project.dbo.covid_infect
Select	inf.continent, inf.location, inf.date, inf.population, 
		vac.new_vaccinations, 
		SUM(CONVERT(bigint,vac.new_vaccinations)) OVER (Partition by inf.location Order by inf.location, inf.date) as vaccinated_total
From portfolio_project.dbo.covid_infect inf
Join portfolio_project.dbo.covid_vaccine vac
	On inf.location = vac.location
	and inf.date = vac.date
Where inf.continent is not null 

Select *
--From "view_percent_population_vaccinated"
From PercentPopulationVaccination

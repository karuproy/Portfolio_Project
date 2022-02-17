## Portfolio_Project contains several of my Past and Current projects


Visit https://public.tableau.com/profile/arup.roy for my Tableau Projects

Please feel free to take a look at my work and give me suggestions

Select location as Country, cast(date as varchar), total_cases as Accumulated_Case, total_deaths as Accumulated_Death, population as Populations
From portfolio_project.dbo.covid_infect
Where cast(date as varchar) like '%Feb  1 2022%'
and continent is not null 
Order by location, date desc

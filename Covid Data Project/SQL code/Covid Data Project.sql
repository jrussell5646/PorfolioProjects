-- Exploring Covid data in SQL
-- Data downloaded from https://ourworldindata.org/covid-deaths
-- Visualizations made with Tableau Public and can be viewed at https://public.tableau.com/app/profile/john1575#!/

Select *
From CovidDeaths
Where continent is not null
Order by 3,4


--Select Data that we are going to be using

Select Location, date, total_cases, new_cases, total_deaths, population
From CovidDeaths
Order by 1,2


-- Looking at Total Cases vs Total Deaths
-- Shows likelihood of dying if you contract covid in the US

Select Location, date, total_cases, total_deaths, (total_deaths/total_cases)*100 as DeathPercentage
From CovidDeaths
Where location='United States'
Order by 1,2


-- Looking at Total Cases vs Population
-- Shows what percentage of population got Covid

Select Location, date, population, total_cases, (total_cases/population)*100 as InfectedPercentage
From CovidDeaths
Where location='United States'
Order by 1,2


-- What countries have the highest infection rate compared to Population

Select Location, population, MAX(total_cases) as HighestInfectionCount, (MAX(total_cases)/population)*100 as InfectedPercentage
From CovidDeaths
Group by location, population
order by 4 desc


-- Showing Countries with Highest Death Count per Population

Select Location, population, MAX(cast(total_deaths as int)) as MaxDeath, (MAX(total_deaths)/population)*100 as DeathPercentage
From CovidDeaths
Where continent is not null
Group by location, population
Order by 3 desc


-- Breaking things down by continent accurately

Select location, MAX(cast(total_deaths as int)) as MaxDeath
From CovidDeaths
where continent is null
		and location not like '%income%'
		and location <> 'World'
		and location <> 'International'
		and location <> 'European Union'
Group by location
Order by 2 desc


-- Let's break things down by continent per Alex

Select continent, MAX(cast(total_deaths as int)) as MaxDeath
From CovidDeaths
Where continent is not null
Group by continent
Order by 2 desc


-- Global Numbers by date

Select date, SUM(new_cases) as total_cases, SUM(cast(new_deaths as int)) as toal_deaths, SUM(cast(new_deaths as int))/SUM(new_cases)*100 as DeathPercentage
From CovidDeaths
--Where location='United States'
Where continent is not null
Group by date
Order by 1,2


-- Global Numbers

Select SUM(new_cases) as total_cases, SUM(cast(new_deaths as int)) as toal_deaths, SUM(cast(new_deaths as int))/SUM(new_cases)*100 as DeathPercentage
From CovidDeaths
--Where location='United States'
Where continent is not null
--group by date
--Order by 1,2


-- Looking at total population vs Vaccinations

Select dea.continent, dea.location, dea.date, dea.population, vac.new_vaccinations, 
	SUM(cast(vac.new_vaccinations as bigint)) Over		(Partition by dea.location 
														Order by dea.location, dea.date) as running_total_vaccinations
From CovidDeaths dea
Join  CovidVaccinations vac
	On dea.location = vac.location
	and	dea.date = vac.date
Where dea.continent is not null
Order by 2,3


-- Use CTE

With PopvsVac (Continent, locatoin, date ,population, new_vaccinations, running_total_vaccinations)
as
(
Select dea.continent, dea.location, dea.date, dea.population, vac.new_vaccinations, 
	SUM(cast(vac.new_vaccinations as bigint)) Over		(Partition by dea.location 
														Order by dea.location, dea.date) as running_total_vaccinations
--, (running_total_vaccinations/population)*100
From CovidDeaths dea
Join  CovidVaccinations vac
	On dea.location = vac.location
	and	dea.date = vac.date
Where dea.continent is not null
--order by 2,3
)
Select *, (running_total_vaccinations/population)*100 as percent_vaccinated_rolling
From PopvsVac


-- Temp Table

DROP table if exists #PercentPopulationVaccinated
Create Table #PercentPopulationVaccinated
(
Continent nvarchar(255),
Location nvarchar(255),
Date datetime,
Population numeric,
New_vaccinations numeric,
running_total_vaccinations numeric
)


Insert into #PercentPopulationVaccinated
Select dea.continent, dea.location, dea.date, dea.population, vac.new_vaccinations, 
	SUM(cast(vac.new_vaccinations as bigint)) Over		(Partition by dea.location 
														Order by dea.location, dea.date) as running_total_vaccinations
--, (running_total_vaccinations/population)*100
From CovidDeaths dea
Join  CovidVaccinations vac
	On dea.location = vac.location
	and	dea.date = vac.date
Where dea.continent is not null
--order by 2,3


Select *, (running_total_vaccinations/population)*100 as percent_vaccinated_rolling
From #PercentPopulationVaccinated


-- Creating Views to store data for later visualizations


-- 1

Create View GlobalNumbers as 
Select SUM(new_cases) as total_cases, SUM(cast(new_deaths as int)) as toal_deaths, SUM(cast(new_deaths as int))/SUM(new_cases)*100 as DeathPercentage
From CovidDeaths
Where continent is not null

-- 2

Create View TotalDeathsPerContinent as
Select location, MAX(cast(total_deaths as int)) as MaxDeath
From CovidDeaths
where continent is null
		and location not like '%income%'
		and location <> 'World'
		and location <> 'International'
		and location <> 'European Union'
Group by location

-- 3

Create View InfectionRatePerCountry as
Select Location, population, MAX(total_cases) as HighestInfectionCount, (MAX(total_cases)/population)*100 as InfectedPercentage
From CovidDeaths
Group by location, population

-- 4

Create View InfectionRatePerCountryPerDate as
Select Location, population, cast(date as date) as Short_date, MAX(total_cases) as HighestInfectionCount, (MAX(total_cases)/population)*100 as InfectedPercentage
From CovidDeaths
Group by location, population, date

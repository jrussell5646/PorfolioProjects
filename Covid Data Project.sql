select *
From AlexPortProject..CovidDeaths
where continent is not null
order by 3,4

--select *
--From AlexPortProject..CovidVaccinations
--order by 3,4

--Select Data that we are going to be using

select Location, date, total_cases, new_cases, total_deaths, population
From AlexPortProject..CovidDeaths
order by 1,2

--Looking at Total Cases vs Total Deaths
-- Shows likelihood of dying if you contract covid in the US

select Location, date, total_cases, total_deaths, (total_deaths/total_cases)*100 as DeathPercentage
From AlexPortProject..CovidDeaths
Where location='United States'
order by 1,2

-- Looking at Total Cases vs Population
-- Shows what percentage of population got Covid

select Location, date, population, total_cases, (total_cases/population)*100 as InfectedPercentage
From AlexPortProject..CovidDeaths
Where location='United States'
order by 1,2

-- What countries have the highest infection rate compared to Population

select Location, population, MAX(total_cases) as HighestInfectionCount, (MAX(total_cases)/population)*100 as InfectedPercentage
From AlexPortProject..CovidDeaths
Group by location, population
order by 4 desc

-- Showing Countries with Highest Death Count per Population

select Location, population, MAX(cast(total_deaths as int)) as MaxDeath, (MAX(total_deaths)/population)*100 as DeathPercentage
From AlexPortProject..CovidDeaths
where continent is not null
Group by location, population
order by 3 desc

-- Let's break things down by continent correctly

select location, MAX(cast(total_deaths as int)) as MaxDeath
From AlexPortProject..CovidDeaths
where continent is null
		and location not like '%income%'
Group by location
order by 2 desc


-- Let's break things down by continent per Alex

select continent, MAX(cast(total_deaths as int)) as MaxDeath
From AlexPortProject..CovidDeaths
where continent is not null
Group by continent
order by 2 desc

-- Global Numbers by date

select date, SUM(new_cases) as total_cases, SUM(cast(new_deaths as int)) as toal_deaths, SUM(cast(new_deaths as int))/SUM(new_cases)*100 as DeathPercentage
From AlexPortProject..CovidDeaths
--Where location='United States'
where continent is not null
group by date
order by 1,2

-- Global Numbers

select SUM(new_cases) as total_cases, SUM(cast(new_deaths as int)) as toal_deaths, SUM(cast(new_deaths as int))/SUM(new_cases)*100 as DeathPercentage
From AlexPortProject..CovidDeaths
--Where location='United States'
where continent is not null
--group by date
order by 1,2

-- Looking at total population vs Vaccinations

Select dea.continent, dea.location, dea.date, dea.population, vac.new_vaccinations, SUM(cast(vac.new_vaccinations as bigint)) Over	(partition by dea.location order by dea.location, dea.date) as running_total_vaccinations

From AlexPortProject..CovidDeaths dea
Join  AlexPortProject..CovidVaccinations vac
	On dea.location = vac.location
	and	dea.date = vac.date
where dea.continent is not null
order by 2,3

-- Use CTE

with PopvsVac (Continent, locatoin, date ,population, new_vaccinations, running_total_vaccinations)
as
(
Select dea.continent, dea.location, dea.date, dea.population, vac.new_vaccinations, SUM(cast(vac.new_vaccinations as bigint)) Over	(partition by dea.location order by dea.location, dea.date) as running_total_vaccinations
--, (running_total_vaccinations/population)*100
From AlexPortProject..CovidDeaths dea
Join  AlexPortProject..CovidVaccinations vac
	On dea.location = vac.location
	and	dea.date = vac.date
where dea.continent is not null
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
Select dea.continent, dea.location, dea.date, dea.population, vac.new_vaccinations, SUM(cast(vac.new_vaccinations as bigint)) Over	(partition by dea.location order by dea.location, dea.date) as running_total_vaccinations
--, (running_total_vaccinations/population)*100
From AlexPortProject..CovidDeaths dea
Join  AlexPortProject..CovidVaccinations vac
	On dea.location = vac.location
	and	dea.date = vac.date
where dea.continent is not null
--order by 2,3

Select *, (running_total_vaccinations/population)*100 as percent_vaccinated_rolling
From #PercentPopulationVaccinated

-- Creating View to store data for later visualizations

Create View PercentPopulationVaccinated as
Select dea.continent, dea.location, dea.date, dea.population, vac.new_vaccinations, SUM(cast(vac.new_vaccinations as bigint)) Over	(partition by dea.location order by dea.location, dea.date) as running_total_vaccinations
--, (running_total_vaccinations/population)*100
From AlexPortProject..CovidDeaths dea
Join  AlexPortProject..CovidVaccinations vac
	On dea.location = vac.location
	and	dea.date = vac.date
where dea.continent is not null
--order by 2,3

Select *
From PercentPopulationVaccinated

-- GDAC Capstone - Bike-share 3/11/2022

-- Data was downloaded from https://divvy-tripdata.s3.amazonaws.com/index.html

-- Previous 12 months (March 2021 - February 2022) of trip data were downloaded.  Each month is speperated into own .csv file

-- .csv files were imported into SQL Server individually

-- Creates new table, TripData_Master, from 03/2021 table

Select * into
TripData_Master
From TripData_2021_03


-- Copies remaining tables into master table. Used multiple times, changed from table as needed

Insert into TripData_Master
Select *
From TripData_2022_02


-- Checking rows in master tables against rows of all monthly tables to verify imports

Select Count(*), left(started_at,7) as Month_
From TripData_Master
Group by left(started_at, 7)
Order by 2

Select Count(*)
From TripData_2022_02


-- Calculates and adds ride time for each trip and creates columns ride_time, day_of_week, month_of_year

Alter Table TripData_Master
Add ride_time as DATEDIFF(minute, started_at, ended_at),
day_of_week as DATENAME(WEEKDAY, started_at),
month_of_year as datename(month, started_at)


-- Checking maximum and minimum values of ride time

Select
	MAX(ride_time) as max_ride_time,
	MIN(ride_time) as min_ride_time
From TripData_Master

-- Deletes negative ride time values from table.  There should be no negative values in this column.

Delete From TripData_Master
Where ride_time <0

-- Creates new table MembersVsCasual to be used in Tableau. Calculates trips per day, average ride time, and max ride time per month per rider type.  Excludes trips longer than 1 day as these are outliers.

Drop Table if exists MemberVsCasual
Select member_casual, day_of_week, month_of_year, count(day_of_week) as Trips_per_day, AVG(ride_time) as avg_ride_time, max(ride_time) as max_ride_time
Into MembersVsCasual
From TripData_Master
Where ride_time < 1440
Group by member_casual, day_of_week, month_of_year
Order by 1, 3, 4 desc

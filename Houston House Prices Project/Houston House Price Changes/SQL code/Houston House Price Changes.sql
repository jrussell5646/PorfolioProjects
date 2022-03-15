-- Data taken from https://www.zillow.com/research/data/ on 3/9/2022.  Used Zillow Home Value Index for All Homes by ZIP Code and Zillow Observed Rent Index for All Homes by ZIP Codes.

-- Data was tranformed from wide to long in Excel with VBA codes and some manual manipulation.  Both data sets were reduced down to TX data only.

-- Visualizations made with Tableau Public and can be viewed at https://public.tableau.com/app/profile/john1575#!/


-- Shows data set for monthly Median Home Values in Texas from Jan 2000 - Jan 2022

Drop table if exists MedianHomePrices1
Select * into MedianHomePrices1
from MedianHomeLong

Alter Table MedianHomePrices1
Alter Column date date

-- Creates table HoustonPricesMoM to show Month over Month change of prices for each zipcode

Drop Table if exists HoustonPricesMoM
Select date, zip, price, round((price - last_month_price) *100/last_month_price,1) as MoM
Into HoustonPricesMoM
From 
	(
Select date, zip, price,
	lag(price) OVER	(Partition by zip
				order by concat(date, zip)) as last_month_price
From MedianHomePrices1
Where metro like '%houston%'
) t;

-- Shows the last year price

Select zip, round(avg(price),2) as avg_price, year(date) as year
From MedianHomePrices1
Where metro like '%houston%'
group by year(date),
		zip
Order by zip, year

-- Creates table HoustonPricesYoY to show Year over Year change of average prices for each zipcode

Drop Table if exists HoustonPricesYoY
Select year, zip, avg_price, round((avg_price - last_year) *100/last_year,1) as YoY
into HoustonPricesYoY
From 
	( 
Select zip, avg_price, year,
	lag(avg_price) OVER (Partition by zip
					order by zip, year) as last_year
From 
	(
Select zip, round(avg(price),2) as avg_price, year(date) as year
From MedianHomePrices1
Where metro like '%houston%'
Group by year(date),
		zip
) t
) t1;
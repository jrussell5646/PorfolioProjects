-- Data taken from https://www.zillow.com/research/data/ on 3/9/2022.  Used Zillow Home Value Index for All Homes by ZIP Code and Zillow Observed Rent Index for All Homes by ZIP Codes.

-- Data was tranformed from wide to long in Excel with VBA codes and some manual manipulation.  Both data sets were reduced down to TX data only.

-- Visualizations made with Tableau Public and can be viewed at https://public.tableau.com/app/profile/jrussell5646


-- Shows data set for monthly Median Home Values in Texas from Jan 2000 - Jan 2022

Select *
From ZillowData..medianhomelong
Order by 1


-- Shows data set for monthly Rental Prices in Texas from Jan 2014 - Jan 2022

Select *
From ZillowData..rentallong
Order by 1


-- Creates new table CommonZips for data that is in both Median Home Values and Rental Prices in Houston Metro Area

Drop Table if exists CommonZips
Select DISTINCT mhl.Zip
into CommonZips
From ZillowData..medianhomelong mhl
Join ZillowData..RentalLong	rll
ON	mhl.zip = Rll.Zip
Where mhl.metro like '%Houston%'


-- Creates new table HoustonHomePrices to be used in Tableau

Drop Table if exists HoustonHomePrices
Select mhl.zip, mhl.city, mhl.Price, LEFT(CAST(date as date), 7) as date, CONCAT(LEFT(CAST(date as date), 7), mhl.zip) as datez
into HoustonHomePrices
From ZillowData..medianhomelong mhl
Right Join CommonZips cz
ON	mhl.zip = cz.zip
Order by 1

Select *
From HoustonHomePrices
--Where city = 'Houston'
Order by Date


-- Creates new table HoustonRentPrices to be used in Tableau

Drop Table if exists HoustonRentPrices
Select rll.date, rll.zip, rll.city, rll.Price as rent, CONCAT(rll.date, rll.zip) as datez
into HoustonRentPrices
From ZillowData..RentalLong rll
Right Join CommonZips cz
ON	rll.zip = cz.zip
Order by 1

Select *
From HoustonRentPrices
--Where city = 'Houston'
Order by Date


-- Creates new Table PriceToRent to be used in Tableau

Drop Table if exists PriceToRent
Select hrp.date, hrp.zip, hrp.city, hrp.rent, hhp.price, hrp.rent*12 as annualrent, ROUND((hrp.rent*12/hhp.price)*100, 2) as ptr
into PriceToRent
From HoustonRentPrices	hrp
Join HoustonHomePrices	hhp
ON	hrp.datez = hhp.datez
Order by date

Select *
From PriceToRent
Where ptr is not null
Order by Date

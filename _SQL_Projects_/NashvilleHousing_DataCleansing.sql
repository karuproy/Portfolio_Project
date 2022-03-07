/*
Cleaning Data in SQL Queries
*/


Select *
From portfolio_project..nashville_housing


--------------------------------------------------------------------------------------------------------------------------


-- Standardize Date Format

Select SaleDate, CONVERT(Date,SaleDate)
From portfolio_project..nashville_housing

Update portfolio_project..nashville_housing
Set SaleDate = CONVERT(Date,SaleDate)


-- Alternate Method
 
 /*
Alter Table portfolio_project..nashville_housing
Add SaleDate_Converted Date;

Update portfolio_project..nashville_housing
Set SaleDate_Converted = CONVERT(Date,SaleDate)

Select SaleDate_Converted, CONVERT(Date,SaleDate)
From portfolio_project..nashville_housing
*/


 --------------------------------------------------------------------------------------------------------------------------


-- Populate Property Address data

Select *
From portfolio_project..nashville_housing
Where PropertyAddress is null
order by ParcelID


Select a.ParcelID, a.PropertyAddress, b.ParcelID, b.PropertyAddress, ISNULL(a.PropertyAddress,b.PropertyAddress)
From portfolio_project..nashville_housing a
Join portfolio_project..nashville_housing b
	on a.ParcelID = b.ParcelID
	AND a.[UniqueID ] <> b.[UniqueID ]
Where a.PropertyAddress is null


Update a
SET PropertyAddress = ISNULL(a.PropertyAddress,b.PropertyAddress)
From portfolio_project..nashville_housing a
JOIN portfolio_project..nashville_housing b
	on a.ParcelID = b.ParcelID
	AND a.[UniqueID ] <> b.[UniqueID ]
Where a.PropertyAddress is null


--------------------------------------------------------------------------------------------------------------------------


-- Breaking out Address into Individual Columns (Address, City, State)

Select PropertyAddress
From portfolio_project..nashville_housing


Select
	SUBSTRING(PropertyAddress, 1, CHARINDEX(',', PropertyAddress) -1 ) as Address,
	SUBSTRING(PropertyAddress, CHARINDEX(',', PropertyAddress) + 1 , LEN(PropertyAddress)) as City
From portfolio_project..nashville_housing



Alter Table portfolio_project..nashville_housing
Add PropertyAddress_splited Nvarchar(255);

Update portfolio_project..nashville_housing
Set PropertyAddress_splited = SUBSTRING(PropertyAddress, 1, CHARINDEX(',', PropertyAddress) -1 )



Alter Table portfolio_project..nashville_housing
Add PropertyCity_splited Nvarchar(255);

Update portfolio_project..nashville_housing
Set PropertyCity_splited = SUBSTRING(PropertyAddress, CHARINDEX(',', PropertyAddress) + 1 , LEN(PropertyAddress))



Select *
From portfolio_project..nashville_housing


--------------------------------------------------------------------------------------------------------------------------


-- Breaking out Address into Individual Columns (Alternate Method)

Select OwnerAddress
From portfolio_project..nashville_housing


Select
	PARSENAME(REPLACE(OwnerAddress, ',', '.') , 3),
	PARSENAME(REPLACE(OwnerAddress, ',', '.') , 2),
	PARSENAME(REPLACE(OwnerAddress, ',', '.') , 1)
From portfolio_project..nashville_housing



Alter Table portfolio_project..nashville_housing
Add OwnerAddress_splited Nvarchar(255);

Update portfolio_project..nashville_housing
Set OwnerAddress_splited = PARSENAME(REPLACE(OwnerAddress, ',', '.') , 3)



Alter Table portfolio_project..nashville_housing
Add OwnerCity_splited Nvarchar(255);

Update portfolio_project..nashville_housing
Set OwnerCity_splited = PARSENAME(REPLACE(OwnerAddress, ',', '.') , 2)



Alter Table portfolio_project..nashville_housing
Add OwnerState_splited Nvarchar(255);

Update portfolio_project..nashville_housing
Set OwnerState_splited = PARSENAME(REPLACE(OwnerAddress, ',', '.') , 1)



Select *
From portfolio_project..nashville_housing


--------------------------------------------------------------------------------------------------------------------------


-- Change Y and N to Yes and No in "Sold as Vacant" field

Select Distinct(SoldAsVacant), Count(SoldAsVacant)
From portfolio_project..nashville_housing
Group by SoldAsVacant
order by 2


Select SoldAsVacant
, CASE When SoldAsVacant = 'Y' THEN 'Yes'
	   When SoldAsVacant = 'N' THEN 'No'
	   ELSE SoldAsVacant
	   END
From portfolio_project..nashville_housing


Update portfolio_project..nashville_housing
SET SoldAsVacant = CASE When SoldAsVacant = 'Y' THEN 'Yes'
	   When SoldAsVacant = 'N' THEN 'No'
	   ELSE SoldAsVacant
	   END


---------------------------------------------------------------------------------------------------------------------------


-- Remove Duplicates

Select *,
	
	ROW_NUMBER() OVER (
	PARTITION BY ParcelID,
				 PropertyAddress_splited,
				 SalePrice,
				 SaleDate_Converted,
				 LegalReference
				 ORDER BY UniqueID) row_num

From portfolio_project..nashville_housing
order by ParcelID




WITH RowNumCTE AS(

	Select *,
		ROW_NUMBER() OVER (
		PARTITION BY ParcelID,
					 PropertyAddress_splited,
					 SalePrice,
					 SaleDate_Converted,
					 LegalReference
					 ORDER BY UniqueID) row_num

	From portfolio_project..nashville_housing
	--order by ParcelID
	)

Select *
From RowNumCTE
Where row_num > 1
Order by PropertyAddress_splited




WITH RowNumCTE AS(
	Select *,
		ROW_NUMBER() OVER (
		PARTITION BY ParcelID,
					 PropertyAddress_splited,
					 SalePrice,
					 SaleDate_Converted,
					 LegalReference
					 ORDER BY UniqueID) row_num

	From portfolio_project..nashville_housing
	--order by ParcelID
	)

DELETE
From RowNumCTE
Where row_num > 1
--Order by PropertyAddress_splited


---------------------------------------------------------------------------------------------------------


-- Delete Unused Columns

Alter Table portfolio_project..nashville_housing
Drop Column OwnerAddress, TaxDistrict, PropertyAddress, SaleDate


Select *
From portfolio_project..nashville_housing


-----------------------------------------------------------------------------------------------
-----------------------------------------------------------------------------------------------


--- Importing Data using OPENROWSET and BULK INSERT	

--  More advanced and looks cooler, but have to configure server appropriately to do correctly
--  Wanted to provide this in case you wanted to try it




--sp_configure 'show advanced options', 1;
--RECONFIGURE;
--GO
--sp_configure 'Ad Hoc Distributed Queries', 1;
--RECONFIGURE;
--GO



--USE PortfolioProject 
--GO 
--EXEC master.dbo.sp_MSset_oledb_prop N'Microsoft.ACE.OLEDB.12.0', N'AllowInProcess', 1 
--GO 
--EXEC master.dbo.sp_MSset_oledb_prop N'Microsoft.ACE.OLEDB.12.0', N'DynamicParameters', 1 
--GO 



---- Using BULK INSERT

--USE PortfolioProject;
--GO
--BULK INSERT nashvilleHousing FROM 'C:\Temp\SQL Server Management Studio\Nashville Housing Data for Data Cleaning Project.csv'
--   WITH (
--      FIELDTERMINATOR = ',',
--      ROWTERMINATOR = '\n'
--);
--GO



---- Using OPENROWSET
--USE PortfolioProject;
--GO
--SELECT * INTO nashvilleHousing
--FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0',
--    'Excel 12.0; Database=C:\Users\alexf\OneDrive\Documents\SQL Server Management Studio\Nashville Housing Data for Data Cleaning Project.csv', [Sheet1$]);
--GO
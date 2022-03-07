/*
Cleaning Data in SQL Queries
*/

Select * 
From portfolio_project..nashvilleHousing


Select *
	into nashville_housing
	from portfolio_project..nashvilleHousing

Select * From nashville_housing


--------------------------------------------------------------------------------------------------------------------------


-- Standardize Date Format

Select SaleDate, CONVERT(Date,SaleDate)
From nashville_housing

Update nashville_housing
Set SaleDate = CONVERT(Date,SaleDate)


-- Alternate Method
 
 /*
Alter Table nashville_housing
Add SaleDate_Converted Date;

Update nashville_housing
Set SaleDate_Converted = CONVERT(Date,SaleDate)

Select SaleDate_Converted, CONVERT(Date,SaleDate)
From nashville_housing
*/


 --------------------------------------------------------------------------------------------------------------------------


-- Populate Property Address data

Select *
From nashville_housing
Where PropertyAddress is null
order by ParcelID


Select a.ParcelID, a.PropertyAddress, b.ParcelID, b.PropertyAddress, ISNULL(a.PropertyAddress,b.PropertyAddress)
From nashville_housing a
Join nashville_housing b
	on a.ParcelID = b.ParcelID
	AND a.[UniqueID ] <> b.[UniqueID ]
Where a.PropertyAddress is null


Update a
SET PropertyAddress = ISNULL(a.PropertyAddress,b.PropertyAddress)
From nashville_housing a
JOIN nashville_housing b
	on a.ParcelID = b.ParcelID
	AND a.[UniqueID ] <> b.[UniqueID ]
Where a.PropertyAddress is null


--------------------------------------------------------------------------------------------------------------------------


-- Breaking out Address into Individual Columns (Address, City, State)

Select PropertyAddress
From nashville_housing


Select
	SUBSTRING(PropertyAddress, 1, CHARINDEX(',', PropertyAddress) -1 ) as Address,
	SUBSTRING(PropertyAddress, CHARINDEX(',', PropertyAddress) + 1 , LEN(PropertyAddress)) as City
From nashville_housing



Alter Table nashville_housing
Add PropertyAddress_splited Nvarchar(255);

Update nashville_housing
Set PropertyAddress_splited = SUBSTRING(PropertyAddress, 1, CHARINDEX(',', PropertyAddress) -1 )



Alter Table nashville_housing
Add PropertyCity_splited Nvarchar(255);

Update nashville_housing
Set PropertyCity_splited = SUBSTRING(PropertyAddress, CHARINDEX(',', PropertyAddress) + 1 , LEN(PropertyAddress))



Select *
From nashville_housing


--------------------------------------------------------------------------------------------------------------------------


-- Breaking out Address into Individual Columns (Alternate Method)

Select OwnerAddress
From nashville_housing


Select
	PARSENAME(REPLACE(OwnerAddress, ',', '.') , 3),
	PARSENAME(REPLACE(OwnerAddress, ',', '.') , 2),
	PARSENAME(REPLACE(OwnerAddress, ',', '.') , 1)
From nashville_housing



Alter Table nashville_housing
Add OwnerAddress_splited Nvarchar(255);

Update nashville_housing
Set OwnerAddress_splited = PARSENAME(REPLACE(OwnerAddress, ',', '.') , 3)



Alter Table nashville_housing
Add OwnerCity_splited Nvarchar(255);

Update nashville_housing
Set OwnerCity_splited = PARSENAME(REPLACE(OwnerAddress, ',', '.') , 2)



Alter Table nashville_housing
Add OwnerState_splited Nvarchar(255);

Update nashville_housing
Set OwnerState_splited = PARSENAME(REPLACE(OwnerAddress, ',', '.') , 1)



Select * From nashville_housing


--------------------------------------------------------------------------------------------------------------------------


-- Change Y and N to Yes and No in "Sold as Vacant" field

Select Distinct(SoldAsVacant), Count(SoldAsVacant)
From nashville_housing
Group by SoldAsVacant
order by 2


Select SoldAsVacant
, CASE When SoldAsVacant = 'Y' THEN 'Yes'
	   When SoldAsVacant = 'N' THEN 'No'
	   ELSE SoldAsVacant
	   END
From nashville_housing


Update nashville_housing
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

From nashville_housing
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

	From nashville_housing
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

	From nashville_housing
	--order by ParcelID
	)

DELETE
From RowNumCTE
Where row_num > 1
--Order by PropertyAddress_splited


---------------------------------------------------------------------------------------------------------


-- Delete Unused Columns

Alter Table nashville_housing
Drop Column OwnerAddress, TaxDistrict, PropertyAddress, SaleDate


Select *
From nashville_housing

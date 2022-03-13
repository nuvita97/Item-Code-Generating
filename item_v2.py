import pandas as pd
import os

# Get realtime datetime
from datetime import datetime
now = datetime.now()
dt_string = now.strftime("%d%m%Y")
dt_rev_string = now.strftime("%Y.%m.%d_%H.%M.%S")

# Read from Request file
filename = 'New Item Request Form.xlsx'
df = pd.read_excel(filename, sheet_name = 'GeneratingForm')

# Convert to 000 in TDG
TDG = df['Tracking Dimension Group'].replace(to_replace = 0, value = '000') 
print(TDG, type(TDG), type(TDG[0]))

# Headers
Import_Header = ['PRODUCTNUMBER', 'PRODUCTNAME', 'ISCATCHWEIGHTPRODUCT', 'ISPRODUCTKIT', 'PRODUCTDIMENSIONGROUPNAME', 'VARIANTCONFIGURATIONTECHNOLOGY', 'ITEMNUMBER',	'BOMUNITSYMBOL',	'DEFAULTLEDGERDIMENSIONDISPLAYVALUE',	'INVENTORYUNITSYMBOL',	'ITEMMODELGROUPID',	'PRODUCTGROUPID',	'PRODUCTIONTYPE', 'PRODUCTSUBTYPE',	'PRODUCTTYPE',	'PURCHASESALESTAXITEMGROUPCODE',	'PURCHASEUNITSYMBOL',	'SALESSALESTAXITEMGROUPCODE',	'SALESUNITSYMBOL',	'STORAGEDIMENSIONGROUPNAME',	'TRACKINGDIMENSIONGROUPNAME', 'PROJECTCATEGORYID', 'PRODUCTCATEGORYHIERARCHYNAME', 'PRODUCTCATEGORYNAME']


# Realtime filename
import_filename = f'import_{dt_rev_string}.xlsx'

# DataFrame
df1 = pd.DataFrame(columns = Import_Header)

df1['PRODUCTNUMBER'] = df['Item Number']
df1['PRODUCTNAME'] = df['Item Name']
df1['ISCATCHWEIGHTPRODUCT'] = df['ISCATCHWEIGHTPRODUCT']
df1['ISPRODUCTKIT'] = df['ISPRODUCTKIT']
df1['PRODUCTDIMENSIONGROUPNAME']= df['Product Dimension Group']
df1['VARIANTCONFIGURATIONTECHNOLOGY'] = df['Variant Config']

df1['ITEMNUMBER'] = df['Item Number']
df1['BOMUNITSYMBOL'] = df['UOM']
df1['DEFAULTLEDGERDIMENSIONDISPLAYVALUE'] = df['Product Line F8']
df1['INVENTORYUNITSYMBOL'] = df['UOM']
df1['ITEMMODELGROUPID'] = df['Item Model Group']
df1['PRODUCTGROUPID'] = df['Item Group System']
df1['PRODUCTIONTYPE'] = df['Production Type']
df1['PRODUCTSUBTYPE'] = df['Product Subtype']
df1['PRODUCTTYPE'] = df['Product Type']
df1['PURCHASESALESTAXITEMGROUPCODE'] = df['Item Sales Tax Group']
df1['PURCHASEUNITSYMBOL'] = df['UOM']
df1['SALESSALESTAXITEMGROUPCODE'] = df['Item Sales Tax Group']
df1['SALESUNITSYMBOL'] = df['UOM']
df1['STORAGEDIMENSIONGROUPNAME'] = df['Storage Dimension Group']
df1['TRACKINGDIMENSIONGROUPNAME'] = TDG

df1['PROJECTCATEGORYID'] = df['Project Categories']
df1['PRODUCTCATEGORYHIERARCHYNAME'] = df['Product Categories']
df1['PRODUCTCATEGORYNAME'] = df['Product Category Name']

# Create Dimensions
for i in range(len(df['Product Line F8'])):
	df1.loc[i,'DEFAULTLEDGERDIMENSIONDISPLAYVALUE'] = '-------' + str(df.loc[i, 'Product Line F8']) + '--' + str(df.loc[i, 'Item Group System']) + '-' + str(df.loc[i, 'Item Number']) + '-'

# Create excel files
df1.to_excel(import_filename, index = False)


#--------------------------------------------------------------------------------------------------------#

# ProductName = df['Item Name']
# ProductSubtype = df['Product Subtype']
# ProductType = df['Product Type']
# ProductDimension = df['Product Dimension Group']
# SDG = df['Storage Dimension Group']
# TDG = df['Tracking Dimension Group']
# VariantConfig = df['Variant Config']
# ICWP = df['ISCATCHWEIGHTPRODUCT']
# IPK = df['ISPRODUCTKIT']
# Unit = df['UOM']
# ItemModelGroup = df['Item Model Group']
# ItemGroup = df['Item Group System']
# ProductionType = df['Production Type']
# SalesTax = df['Item Sales Tax Group']
# F8 = df['Product Line F8']
# ItemNumber = df['Item Number']
# ProjectCat = df['Project Categories']


# df1['PRODUCTDIMENSIONGROUPNAME']= ProductDimension
# # df1['PRODUCTSEARCHNAME'] = ProductSearchName
# df1['PRODUCTSUBTYPE'] = ProductSubtype
# df1['PRODUCTTYPE'] = ProductType
# df1['STORAGEDIMENSIONGROUPNAME'] = SDG
# df1['TRACKINGDIMENSIONGROUPNAME']= TDG
# df1['VARIANTCONFIGURATIONTECHNOLOGY'] = VariantConfig

# # DataFrame 2 - ReleasedProductsV2_1

# df2 = pd.DataFrame(columns = ReleasedProductsV2Header)
# df2['ITEMNUMBER'] = ItemNumber
# df2['PRODUCTNUMBER'] = ItemNumber
# df2['BOMUNITSYMBOL'] = Unit
# df2['INVENTORYUNITSYMBOL'] = Unit
# df2['ITEMMODELGROUPID'] = ItemModelGroup
# df2['PRODUCTGROUPID'] = ItemGroup
# df2['PRODUCTIONTYPE'] = ProductionType
# # df2['PRODUCTSEARCHNAME'] = ProductSearchName
# df2['PRODUCTSUBTYPE'] = ProductSubtype
# df2['PRODUCTTYPE'] = ProductType
# df2['PURCHASESALESTAXITEMGROUPCODE'] = SalesTax
# df2['PURCHASEUNITSYMBOL'] = Unit
# df2['SALESSALESTAXITEMGROUPCODE'] = SalesTax
# df2['SALESUNITSYMBOL'] = Unit
# # df2['SEARCHNAME'] = ProductSearchName
# df2['STORAGEDIMENSIONGROUPNAME'] = SDG
# df2['TRACKINGDIMENSIONGROUPNAME']= TDG
# df2['DEFAULTLEDGERDIMENSIONDISPLAYVALUE'] = F8

# for i in range(len(F8)):
# 	df2.loc[i,'DEFAULTLEDGERDIMENSIONDISPLAYVALUE'] = '-------' + str(df2.loc[i,'DEFAULTLEDGERDIMENSIONDISPLAYVALUE']) + '----'
# # print(df2.iloc[5]['DEFAULTLEDGERDIMENSIONDISPLAYVALUE'])

# # DataFrame 3 - ReleasedProductsV2_2

# df3 = pd.DataFrame(columns = ReleasedProductsV2Header_2)
# df3['ITEMNUMBER'] = ItemNumber
# df3['PRODUCTNUMBER'] = ItemNumber
# # df3['BOMUNITSYMBOL'] = Unit
# # df3['INVENTORYUNITSYMBOL'] = Unit
# # df3['ITEMMODELGROUPID'] = ItemModelGroup
# # df3['PRODUCTGROUPID'] = ItemGroup
# # df3['PRODUCTIONTYPE'] = ProductionType
# # # df3['PRODUCTSEARCHNAME'] = ProductSearchName
# # df3['PRODUCTSUBTYPE'] = ProductSubtype
# # df3['PRODUCTTYPE'] = ProductType
# # df3['PURCHASESALESTAXITEMGROUPCODE'] = SalesTax
# # df3['PURCHASEUNITSYMBOL'] = Unit
# # df3['SALESSALESTAXITEMGROUPCODE'] = SalesTax
# # df3['SALESUNITSYMBOL'] = Unit
# # # df3['SEARCHNAME'] = ProductSearchName
# # df3['STORAGEDIMENSIONGROUPNAME'] = SDG
# # df3['TRACKINGDIMENSIONGROUPNAME']= TDG
# df3['DEFAULTLEDGERDIMENSIONDISPLAYVALUE'] = F8
# df3['PROJECTCATEGORYID'] = ProjectCat

# for i in range(len(F8)):
# 	df3.loc[i,'DEFAULTLEDGERDIMENSIONDISPLAYVALUE'] = '-------' + str(df.loc[i, 'Product Line F8']) + '--' + str(df.loc[i, 'Item Group System']) + '-' + str(df.loc[i, 'Item Number']) + '-'
# # print(df3.iloc[5]['DEFAULTLEDGERDIMENSIONDISPLAYVALUE'])

# # DataFrame 4 - ProductCategoryAssignments
# df4 = pd.DataFrame(columns = ProductCategoryAssignments)
# df4['PRODUCTNUMBER'] = ItemNumber

# # Create folder
# os.mkdir(dt_rev_string)
# os.chdir(dt_rev_string)





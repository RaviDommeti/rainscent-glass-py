import pandas as pds
vendor_cp =pds.read_excel('SAMPLE Account 1.xlsx',sheet_name = 'CP')
vendor_hsil =pds.read_excel('SAMPLE Account 1.xlsx',sheet_name = 'HSIL')
vendor_liverleaf =pds.read_excel('SAMPLE Account 1.xlsx',sheet_name = 'LIVER LEAF',skiprows=1)
account_file = pds.read_excel('ACCOUNT_PAYMENT_BANK.xlsx')

# print("\n\tVendor CP")
# print(vendor_cp.head())
# print("\n\tVendor HSIL")
# print(vendor_hsil.head())
# print("\n\tVendor LIVER LEAF")
# print(vendor_liverleaf.head())


#clean HSIL containing empty Columns
vendor_hsil = vendor_hsil.dropna(axis=1,how='all')
#vendor_hsil.to_excel('Clean Col HSIL.xlsx')

#Clean Liver leaf with containing empty columns
vendor_liverleaf = vendor_liverleaf.dropna(axis=1,how='all')
#vendor_liverleaf.to_excel('Clean Col Liver Leaf.xlsx')

#Clean Liver Leaf containing unnamed rows
vendor_liverleaf = vendor_liverleaf.dropna(how='all')
#vendor_liverleaf.to_excel('Clean Row Liver Leaf.xlsx')

# #for deleting empty rows
# clean_row_cp = vendor_cp.dropna(how='all')
# clean_row_hsil = vendor_hsil.dropna(how='all')
# clean_row_liverleaf = vendor_liverleaf.dropna(how='all')
#
# clean_row_hsil.to_excel('Clean Row HSIL.xlsx')
# clean_row_liverleaf.to_excel('Clean Row Liver Leaf.xlsx')
#
# #for deleting empty columns
# vendor_cp = clean_row_cp.dropna(axis=1,how='all')
# vendor_hsil = clean_row_hsil.dropna(axis=1,how='all')
# vendor_liverleaf = clean_row_liverleaf.dropna(axis=1,how='all')
#
# vendor_hsil.to_excel('Clean Col HSIL.xlsx')
# vendor_liverleaf.to_excel('Clean Col Liver Leaf.xlsx')

#print (vendor_cp)
vendor_cp['Source'] = 'CP'
vendor_hsil['Source'] = 'HSIL'
vendor_liverleaf['Source'] = 'Liver Leaf'

#Sorting based on columns in ascending order
# sorted_cp = vendor_cp.sort_index(axis=1)
# sorted_hsil = vendor_hsil.sort_index(axis=1)
# sorted_liverleaf = vendor_liverleaf.sort_index(axis=1)
result_output_file = pds.concat([vendor_hsil,vendor_liverleaf])
#Consolidating of HSIL & Liver Leaf from Sample Account 1.xslx
# result_output_file = pds.concat([sorted_hsil,sorted_liverleaf])
# Renaming column 'Truck No' to upper case & add space for comparison as one file has upper case and other has lower case
result_output_file.rename(columns = {'Truck No' : 'TRUCK NO  '},inplace = True)
#print(list(result_output_file))

# Comparing Output File(Consolidation of HSIL & Liver Leaf) with ACCOUNT_PAYMENT_BANK
# with TRUCK NO as key. Please note the space in Column name TRUCK NO
result_final = result_output_file.merge(account_file,how="inner",on = 'TRUCK NO  ')
#pd.merge(a, b, left_on = 'a_col', right_on = 'b_col', how = 'left')
#result_final = pds.merge(account_file,result_output_file, left_on = 'TRUCK NO', right_on = 'Truck No', how = 'inner')
#Contate HSIL and LIVER Leaf
#result_final = pds.concat([sorted_hsil,sorted_liverleaf])
#result_final = pds.concat([result_final,sorted_cp])

#List of column names
print("\n\n\t\t Output File")
print(list(result_output_file))
print("\n\n\t\t Account File")
print(list(account_file))
print("\n\n\t\t Final File")
print(list(result_final))

# print("\n\tSorted Vendor CP")
# print(sorted_cp.head())
# print("\n\tSorted Vendor HSIL")
# print(sorted_hsil.head())
# print("\n\tSorted Vendor LIVER LEAF")
# print(sorted_liverleaf.head())
# print("\n\tVendor FINAL Table")
# print(result_final.head(10))

#Create output excel file
result_output_file.to_excel('Output File.xlsx')
result_final.to_excel('Account Summary.xlsx')

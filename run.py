import openpyxl
from openpyxl.styles import Font, NamedStyle

# Find the sheet section with the
# end_table = subtract - 1
# get the value above
# get the next table name list and the coordinate of the next blank cell above.
# the end is the coordinate of the second item of extract_table_name_list minus one.
# size = row_number - subtract
# print('size of table: {}'.format(str(size)))
# recognize the begin and the and of the tables.
# read the cell and consult the extract_table_name_list
# copy value and coondinate
# find the last cell of the table
# Define the type of the table.
# if the line is blank and is only one row
# TIP: we already have the tables names position
# just get the blank line above to find the end of the last table.
# verify if is only one row
# and above nd bottom line is not in
# table_name_used o r extract_table_name_list
# get the coordinate


filename = r"D:\Arquivos HD\Projetos HD\SD Labs\transformer\SPSS_output_scrapper-master\WEIGHTED_customer_database.xlsx"

# Import your dataset, for example:
wb = openpyxl.load_workbook(filename)

sheet=wb["Sheet1"]
predefined_table_name_list = ['$Var_set*agecat Crosstabulation', '$Var_set Frequencies', 'Notes', 'Case Processing Summary', 'Gender * Age category Crosstabulation', 'Union member * Age category Crosstabulation', 'Retired * Age category Crosstabulation', 'Marital status * Age category Crosstabulation']
for i in sheet['A1':'A161']:#section
    for n in i:
        if n.value in predefined_table_name_list:
            print(n.coordinate, n.value)

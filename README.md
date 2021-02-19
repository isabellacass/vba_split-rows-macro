# vba_split-rows-macro

# this macro was written to clean a data set of anonymized crime data for the city of Salinas, California, for a multiple regression analysis in R 
# each crime was recorded on a single row, even if there were multiple suspects
# for crimes with multiple suspects, the suspects' biographical data were grouped in a single cell
# to perform a regression using the suspects' age, the age data for crimes with multiple suspects needed to be split into multiple rows so that each row contained only one suspect age
# this macro split rows with multiple age values into multiple rows for each suspect age value
# a multiple regression using suspect age was then run on the cleaned data in R

# vba_split-rows-macro

# this macro was written to clean a data set of anonymized crime data for the city of Salinas, California, for a multiple regression analysis in R 
# each crime was recorded on a single row, even if there were multiple suspects
# for crimes with multiple suspects, the suspects' biographical data were grouped in a single cell
# to perform a regression using the suspects' age, the age data for crimes with multiple suspects needed to be split so that each age was in a separate cell
# this macro split cells with multiple age values into separate cells for the age of each suspect
# a multiple regression using suspect age was then run on the cleaned data in R

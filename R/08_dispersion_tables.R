library(readxl)
library(openxlsx) #devtools::install_github("awalker89/openxlsx")
library(tidyverse)
library(janitor)


# Create new workbook
#new_rass <- createWorkbook()


# Load default RASS template
disp <- readWorkbook("Test_R_RASS-Copy.xlsx", sheet = 10, colNames = F, rowNames = F)

#------------------------------#
# Create Dispers Tables SHEET
#------------------------------#

# Add a sheet
addWorksheet(new_rass, "Disp Tables")

# Create content
new <- tibble()

new[1, 1] <- ""

new[ , 2:34] <- rep("", 33)

new[1, 1:3] <- c("No inputs on this page", " ", " ")

new[1 , 4:34] <- ""

new[2:601, 1:34] <- disp[2:601, 1:34]


# Add text
writeData(new_rass, "Disp Tables", x = new, rowNames = F, colNames = F)


# SAVE
#saveWorkbook(new_rass, file = "08_Dispersion_tables.xlsx", overwrite = T)

#

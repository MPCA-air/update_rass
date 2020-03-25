library(readxl)
library(openxlsx) #devtools::install_github("awalker89/openxlsx")
library(tidyverse)

# Create new workbook
new_rass <- createWorkbook()

# Add sheets
source("R/01_Create_emissions.R")
source("R/07_Stack_parameters.R")
source("R/06_Create_summary.R")
source("R/03_Create_risks.R")
source("R/02_Create_concentrations.R")
addWorksheet(new_rass, "Tox Values")
source("R/04_Create_MPSFs.R")
source("R/05_Create_early_life_adjust.R")
source("R/08_Dispersion_tables.R")

# Save
saveWorkbook(new_rass, file = paste0(Sys.Date(), "_Auto_Rass.xlsx"), overwrite = T)


#

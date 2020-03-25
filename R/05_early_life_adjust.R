library(readxl)
library(openxlsx) #devtools::install_github("awalker89/openxlsx")
library(tidyverse)
library(janitor)


# Create new workbook
#new_rass <- createWorkbook()


# New pollutant list
copc <- read_csv("updated_pollutant_list.csv")


# Add if Multi-media pollutant
early <- read_csv("https://raw.githubusercontent.com/MPCA-air/health-values/master/early_life_adjustment_ADAFs.csv") %>%
         clean_names() %>%
         rename(CAS = cas)


copc <- left_join(select(early, -pollutant, -unit_risk_ug_m3_1, -x10_5_cancer_based_air_conc_ug_m3),
                  select(copc, CAS, `Chemical Name`))


#------------------------------------#
# Create Early Life Adjustment sheet
# ADAFs -> Age Dependent Adjustment Factor
# MDH recommends 1.6X cancer risk
#-----------------------------------#

# Add a sheet
addWorksheet(new_rass, "Early Life Adjust")

# Create content
new <- tibble()

new[1:3, 1] <- c("=Summary!A1", "No inputs on this page", " ")

new[ , 2:5] <- ""

new[4, 1:5] <- c("CAS# or MPCA#",
                 "Chemical Name",
                 "Early life adjustment needed?",
                 "Toxicity value source",
                 "Notes")

# Add Early life table
for(i in 1:nrow(copc)) {
  new[4+i, 1:5] <- copc[i, c(1,5,3,2,4)]
}

# Add text
writeData(new_rass, "Early Life Adjust", x = new, rowNames = F, colNames = F)

# Add version
writeFormula(new_rass,
             "Early Life Adjust", x = "=Summary!A1", startRow = 1, startCol = 1)

# Save
saveWorkbook(new_rass, file = "05_Early_Life_sheet.xlsx", overwrite = T)

#


library(readxl)
library(openxlsx) #devtools::install_github("awalker89/openxlsx")
library(tidyverse)
library(janitor)


# Create new workbook
#new_rass <- createWorkbook()


# New pollutant list
copc <- read_csv("updated_pollutant_list.csv")


# Add if Multi-media pollutant
multi_media <- read_csv("https://raw.githubusercontent.com/MPCA-air/health-values/master/multi_pathway_risk_factors.csv") %>%
               clean_names() %>%
               rename(CAS = cas)

multi_media[is.na(multi_media)] <- 0


copc <- left_join(select(copc, CAS, `Chemical Name`),
                  select(multi_media, -pollutant_mpsf))


#------------------------------#
# Create MPS Factors SHEET
#------------------------------#

# Add a sheet
addWorksheet(new_rass, "MPS Factors")

# Create content
new <- tibble()

new[1:3, 1] <- c("=Summary!A1", "No inputs on this page", " ")

new[, 2:8] <- ""

new[4, 1:8] <- c("CAS# or MPCA#", "Chemical Name",
                 "Resident Noncancer", "Resident Cancer",
                 "Urban Gardener Noncancer", "Urban Gardener Cancer",
                 "Farmer Noncancer",	"Farmer Cancer")

# Add Multi-media factors
for(i in 1:nrow(copc)) {
  new[4+i, 1:8] <- copc[i, 1:8]
}

# Add text
writeData(new_rass, "MPS Factors", x = new, rowNames = F, colNames = F)

# Add version
writeFormula(new_rass,
             "MPS Factors", x = "=Summary!A1", startRow = 1, startCol = 1)


# Save
saveWorkbook(new_rass, file = "04_MPS_Factors_sheet.xlsx", overwrite = T)
#

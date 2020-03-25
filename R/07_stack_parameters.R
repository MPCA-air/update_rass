library(readxl)
library(openxlsx) #devtools::install_github("awalker89/openxlsx")
library(tidyverse)
library(janitor)


# Create new workbook
#new_rass <- createWorkbook()


# Load default RASS template
stacks <- readWorkbook("Test_R_RASS-Copy.xlsx", sheet = 4, colNames = F, rowNames = F) %>%
          .[9:14, ]

stacks[1, 3:27] <- gsub("k", "k ", stacks[1, 3:27])

stacks[1:3, 1] <- gsub("m", "meters", stacks[1:3, 1])

#------------------------------#
# Create Dispersion SHEET
#------------------------------#

# Add a sheet
addWorksheet(new_rass, "Dispersion")

# Create content
new <- tibble()

new[1:3, 1] <- c("=Summary!A1",
                 "Inputs should be made in yellow cells",
                 "*AERMOD or other refined air dispersion modeling can be run separately and results entered in place of the default screening values.")

new[ , 2:27] <- ""

new[5:10, 1:27] <- stacks[1:6, 1:27]

new[8:10, 2] <- "auto-lookup or enter manually"

# Add text
writeData(new_rass, "Dispersion", x = new, rowNames = F, colNames = F)

# Add version
writeFormula(new_rass, "Dispersion", x = "=Summary!A1", startRow = 1, startCol = 1)

# Add formula for automatic-lookup
for(stack in 1:25) {

  # Stack column
  stack_col <- c(LETTERS[-c(1:2)], "AA")[stack]

  # 1-hr dispersion
  writeFormula(new_rass, "Dispersion",
               x = paste0('=IF(OR(',
                          stack_col, '6="",',
                          stack_col, '7=""),"",HLOOKUP(',
                          stack_col, '7,$\'Disp Tables\'.$E$2:$AH$101,',
                          stack_col, '6+1)))'),
               startRow = 8,
               startCol = 2+stack)

  # 24-hr dispersion
  writeFormula(new_rass, "Dispersion",
               x = paste0('=IF(OR(',
                          stack_col, '6="",',
                          stack_col, '7=""),"",HLOOKUP(',
                          stack_col, '7,$\'Disp Tables\'.$E$302:$AH$401,',
                          stack_col, '6+1)))'),
               startRow = 9,
               startCol = 2+stack)

  # Annual dispersion
  writeFormula(new_rass, "Dispersion",
               x = paste0('=IF(OR(',
                          stack_col, '6="",',
                          stack_col, '7=""),"",HLOOKUP(',
                          stack_col, '7,$\'Disp Tables\'.$E$502:$AH$601,',
                          stack_col, '6+1)))'),
               startRow = 10,
               startCol = 2+stack)
  }

# SAVE
#saveWorkbook(new_rass, file = "07_Dispersion_sheet.xlsx", overwrite = T)

#

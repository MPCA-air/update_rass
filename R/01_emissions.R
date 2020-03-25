library(readxl)
library(openxlsx) #devtools::install_github("awalker89/openxlsx")
library(tidyverse)


# Create new workbook
#new_rass <- createWorkbook()


# Load default RASS template
emit <- readWorkbook("Test_R_RASS-Copy.xlsx", sheet = 3)


# Load full pollutant table
copc <- read_csv("temp_updated_tox_values_refs.csv")

# Sort by name
copc  <- arrange(copc, chemical_name)


# Switch to Excel column names
names(copc) <- gsub("_", " ", names(copc)) %>%
               str_to_title() %>%
               gsub("Non Cancer", "Non-Cancer", .) %>%
               gsub("Sub Chronic", "Sub-Chronic", .)

# Round digits
copc <- copc %>%
        rowwise() %>%
        rename(CAS = Cas) %>%
        mutate(CAS                                    = str_trim(CAS),
               `Chemical Name`                        = str_trim(`Chemical Name`),
               `Acute Air Concentration`              = signif(as.numeric(`Acute Air Concentration`), digits = 3),
               `Cancer Based Air Concentration`       = signif(as.numeric(`Cancer Based Air Concentration`), digits = 3),
               `Non-Cancer Chronic Air Concentration` = signif(as.numeric(`Non-Cancer Chronic Air Concentration`), digits = 3),
               `Sub-Chronic Air Concentration`        = signif(as.numeric(`Sub-Chronic Air Concentration`), digits = 3)) %>%
        filter(!is.na(CAS), !is.na(`Chemical Name`))

# Replace NA's with spaces
copc <- copc %>%
        mutate_all(as.character) %>%
        mutate(`Acute Air Concentration`              = ifelse(is.na(`Acute Air Concentration`), " ", `Acute Air Concentration`),
               `Cancer Based Air Concentration`       = ifelse(is.na(`Cancer Based Air Concentration`), " ", `Cancer Based Air Concentration`),
               `Non-Cancer Chronic Air Concentration` = ifelse(is.na(`Non-Cancer Chronic Air Concentration`), " ", `Non-Cancer Chronic Air Concentration`),
               `Acute Air Reference`                  = ifelse(is.na(`Acute Air Reference`), " ", `Acute Air Reference`),
               `Cancer Based Air Reference`           = ifelse(is.na(`Cancer Based Air Reference`), " ", `Cancer Based Air Reference`),
               `Non-Cancer Chronic Air Reference`     = ifelse(is.na(`Non-Cancer Chronic Air Reference`), " ", `Non-Cancer Chronic Air Reference`))

rownames(copc) <- NULL


# Drop ethanol pollutants??
#copc <- copc[1:(grep("Acetic Acid", copc$`Chemical Name`)-1), ]

write_csv(copc, "updated_pollutant_list.csv")


#------------------------------#
# Create emissions SHEET
#------------------------------#
new <- tibble()

new[1, 1] <- "=Summary!A1"

new[1:6, 2] <- c("Screening Date: ", "AQ Facility ID: ", "AQ File: ",
                 "Facility Name: ", "Facility Location: ", "SIC Code: ")

new[, 3:54] <- ""

new[8, 1:(4+25+25)] <- c("CAS# or MPCA#", "Chemical Name", "", "Total Annual Emissions",
                         rbind(c("Stack #1",	"Stack #2",		"Stack #3",		"Stack #4",		"Stack #5",		"Stack #6",		"Stack #7",		"Stack #8",		"Stack #9",		"Stack #10",		"Stack #11",		"Stack #12",		"Stack #13",		"Stack #14",		"Stack #15",		"Stack #16",		"Stack #17",		"Stack #18",		"Stack #19",		"Stack #20",		"Stack #21",		"Stack #22",		"Stack #23",
                        "Stack #24", "Stack #25"), c(" ")) %>% as.vector())

new[10, 1:(4+25+25)] <- c("", "", "", "(tons)", rep(c("Hourly Emissions (lb/hr)", "Annual Emissions (tpy)"), 25))

# Add pollutant rows
for(i in 1:nrow(copc)) {
        new[10+i, 1] <- copc[i, "CAS"]
        new[10+i, 2] <- copc[i, "Chemical Name"]
}

new_emit <- new


# Add a sheet
addWorksheet(new_rass, "Emissions (start here)")

# Add text
writeData(new_rass, "Emissions (start here)", x = new_emit, rowNames = F, colNames = F)


# Add version & formulas
writeFormula(new_rass,
             "Emissions (start here)", x = "=Summary!A1", startRow = 1, startCol = 1)

# Emissions totals
for(i in 1:nrow(copc)) {
        writeFormula(new_rass,
                     "Emissions (start here)",
                     x        = gsub("5", 10+i, "=F5+H5+J5+L5+N5+P5+R5+T5+V5+X5+Z5+AB5+AD5+AF5+AH5+AJ5+AL5+AN5+AP5+AR5+AT5+AV5+AX5+AZ5+BB5"),
                     startRow = 10+i,
                     startCol = 4)
}

# Save
#saveWorkbook(new_rass, file = "01_Emissions_sheet.xlsx", overwrite = T)

#

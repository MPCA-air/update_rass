library(readxl)
library(openxlsx) #devtools::install_github("awalker89/openxlsx")
library(tidyverse)
library(janitor)


# Create new workbook
#new_rass <- createWorkbook()

# New pollutant list
copc <- read_csv("updated_pollutant_list.csv")

# Check health endpoints
endpoints <- read_csv("https://raw.githubusercontent.com/MPCA-air/health-values/master/pollutant_endpoints.csv") %>%
             clean_names() %>%
             rename(CAS = cas) %>%
             select(-pollutant_endpt)

copc <- left_join(copc, endpoints)


#------------------------------#
# Create Summary
#------------------------------#

# Add a sheet
addWorksheet(new_rass, "Summary")

# Create text content
new <- tibble()

new[1:4, 1] <- c("RASS version: 2020-03", "No inputs on this page", " ", "Air Toxics Screening")

new[ , 2:16] <- ""

new[5, 1:16] <- c("Total Inhalation Hazard Indices and Cancer Risks", rep("", 3),
                  "Total Indirect Pathway Hazard Indices and Cancer Risks", rep("", 5),
                  "Total Multipathway Hazard Indices and Cancer Risks", rep("", 5))

new[6, ] <- c("Acute", "Subchronic Noncancer", "Chronic Noncancer",	"Cancer",
              rep(c("Farmer Noncancer",	"Farmer Cancer",	"Urban Gardener Noncancer",
                    "Urban Gardener Cancer", "Resident Noncancer", "Resident Cancer"), 2))

# Add Guidance level
new[8, c(1:3, seq(5, 15, 2))] <- 1.49

new[8, seq(4, 16, 2)] <- 0.00001499

# Total row labels
new[7:10, 17] <- c("<-- Rounded value for reporting",
                   "<-- Guidance level",
                   "<-- OK or REFINE?",
                   "<-- Calculated value for transparency and further calculation")

# Endpoint table
endpoints_names <- c("Auditory",
                     "Blood / Hematological",
                     "Bone / Teeth",
                     "Cardiovascular",
                     "Digestive",
                     "Eyes",
                     "Kidney",
                     "Liver",
                     "Neurological",
                     "Reproductive / Developmental / Endocrine",
                     "Respiratory",
                     "Skin")

codes <- c("Auditory", "Blood", "Bone/Teeth", "Cardio", "Digest", "Eyes",
           "Kidney", "Liver", "Neuro", "Repro", "Resp", "Skin")

endpoints_names <- tibble(endpoints = endpoints_names, codes = codes)

new[12:13, 1] <- c("Air Toxics Endpoint Refinement", "Total Inhalation Hazard Indices")
new[14, 1:5]  <- c("", rep(0.99, 3), "<-- Guidance level")
new[15, 5] <- "<-- OK or REFINE?"
new[16, 1:4]  <- c("Endpoint", "Acute", "Subchronic Noncancer", "Chronic Noncancer")

new[1:(16+length(endpoints_names$endpoints)), 1] <- c(unlist(new[1:16, 1]), endpoints_names$endpoints)

new[18+nrow(endpoints_names), 1] <- "Some pollutants have more than one endpoint and are included in multiple endpoint totals."


# Ceiling values table
ceilings <- c("Benzene",
              "Bromopropane, 1-",
              "Butadiene, 1,3-",
              "Carbon disulfide",
              "Cellosolve Acetate (ethylene glycol monoethyl ether acetate)",
              "Chloroform",
              "Ethoxyethanol, 2- (ethylene glycol monoethyl ether)",
              "Ethyl benzene",
              "Ethyl chloride (Chloroethane)",
              "Methoxyethanol, 2- (ethylene glycol monomethyl ether EGME)",
              "Trichloroethylene",
              "Arsenic",
              "Carbon tetrachloride",
              "Mercury (elemental)",
              "Propylene oxide")

ceilings <- select(copc, CAS, `Chemical Name`) %>%
            filter(`Chemical Name` %in% ceilings) %>%
            group_by(`Chemical Name`) %>%
            slice(1) %>%
            arrange(`Chemical Name`) %>%
            rowwise() %>%
            mutate(`Chemical Name` = str_split(`Chemical Name`, "[(]ethylene g")[[1]][1],
                   `Chemical Name` = ifelse(`Chemical Name` == "Mercury (elemental)", "Mercury", `Chemical Name`),
                   metal           = `Chemical Name` %in% c("Mercury", "Arsenic"))

new[12, 9] <- "Ceiling Values Exceeded?"

new[1:(12+nrow(ceilings)), 9] <-  c(unlist(new[1:12, 9]), ceilings$"Chemical Name")

new[1:(12+nrow(ceilings)), 10] <-  c(unlist(new[1:12, 10]), ceilings$CAS)


# Add text
writeData(new_rass, "Summary", x = new, rowNames = F, colNames = F)


#---------------------------#
# Formulas
#---------------------------#

# Add rounded total
for(i in 1:16) {

  letter <- LETTERS[-c(1,2)][i]

  writeFormula(new_rass,
               "Summary",
               x        = gsub("H", letter, "='Risk Calcs'!H20"),
               startRow = 7,
               startCol = i)
}


# Test whether Risk Total > Guidance level

## 1-hr Acute
writeFormula(new_rass,
               "Summary",
               x        = '=IF(B15="OK",IF(A7>A8,"REFINE","OK"),"REFINE ENDPOINTS")',
               startRow = 9,
               startCol = 1)


## Subchronic
writeFormula(new_rass,
               "Summary",
               x        = '=IF(C15="OK",IF(B7>B8,"REFINE","OK"),"REFINE ENDPOINTS")',
               startRow = 9,
               startCol = 2)


## Chronic Inhale

# Non-cancer
writeFormula(new_rass,
               "Summary",
               x        = '=IF(D15="OK",IF(C7>C8,"REFINE","OK"),"REFINE ENDPOINTS")',
               startRow = 9,
               startCol = 3)

# Cancer
writeFormula(new_rass,
             "Summary",
             x        = '=IF(D7>D8,"REFINE","OK")',
             startRow = 9,
             startCol = 4)

## Chronic Media
for(i in 5:16) {

  letter <- LETTERS[i]

  writeFormula(new_rass,
               "Summary",
               x        = paste0('=IF(', letter, '7>', letter, '8,"REFINE","OK")'),
  startRow = 9,
  startCol = i)

}

# Add 2-digit total
for(i in 1:16) {

  letter <- LETTERS[i]

  writeFormula(new_rass,
               "Summary",
               x        = gsub("H", letter, "=H7"),
               startRow = 10,
               startCol = i)
}


# Endpoint risk sums
for(i in 1:nrow(endpoints_names)) {


  # 1-hr Acute
  pollutant_rows <- 6 + grep(paste0(endpoints_names[i, ]$codes, "|Systemic"),
                             endpoints$acute_toxic_endpoints)

  # Paste Zero if no pollutants have that endpoint
  if(length(pollutant_rows) < 1) {
    pollutant_rows <- 0
  } else pollutant_rows <- paste0("'Risk Calcs'!C", pollutant_rows, collapse = ",")

  writeFormula(new_rass,
               "Summary",
               x        = paste0("=SUM(", pollutant_rows, ")"),
               startRow = 16+i,
               startCol = 2)

  # Subchronic
  pollutant_rows <- 6 + grep(paste0(endpoints_names[i, ]$codes, "|Systemic"),
                             endpoints$subchronic_toxic_endpoints)

  if(length(pollutant_rows) < 1) {
    pollutant_rows <- 0
  } else pollutant_rows <- paste0("'Risk Calcs'!D", pollutant_rows, collapse = ",")

  writeFormula(new_rass,
               "Summary",
               x        = paste0("=SUM(", pollutant_rows, ")"),
               startRow = 16+i,
               startCol = 3)

  # Chronic
  pollutant_rows <- 6 + grep(paste0(endpoints_names[i, ]$codes, "|Systemic"),
                             endpoints$chronic_noncancer_endpoints)

  if(length(pollutant_rows) < 1) {
    pollutant_rows <- 0
  } else pollutant_rows <- paste0("'Risk Calcs'!E", pollutant_rows, collapse = ",")


  writeFormula(new_rass,
               "Summary",
               x        = paste0("=SUM(", pollutant_rows, ")"),
               startRow = 16+i,
               startCol = 4)

}

# Test whether Endpoint Risks > Guidance level
## 1-hr Acute
writeFormula(new_rass,
             "Summary",
             x        = '=IF(MAX(B17:B28)>B14,"REFINE","OK")',
             startRow = 15,
             startCol = 2)

## Subchronic
writeFormula(new_rass,
             "Summary",
             x        = '=IF(MAX(C17:C28)>C14,"REFINE","OK")',
             startRow = 15,
             startCol = 3)

## Chronic
writeFormula(new_rass,
             "Summary",
             x        = '=IF(MAX(D17:D28)>D14,"REFINE","OK")',
             startRow = 15,
             startCol = 4)


# Check pollutants with child ceiling values
for(i in 1:nrow(ceilings)) {

  pollutant_rows <- 6 + grep(ceilings$CAS[i], copc$CAS)

  if(ceilings$metal[i]) {
    pollutant_rows <- 6 + grep(ceilings$`Chemical Name`[i], copc$`Chemical Name`)
  }

  writeFormula(new_rass,
               "Summary",
               x        = paste0("=IF(MAX(",
                                 paste0("'Risk Calcs'!E", pollutant_rows, collapse = ","),
                                 ")>=1,", '"YES","NO")'),
               startRow = 12+i,
               startCol = 11)

}

# Save
#saveWorkbook(new_rass, file = "06_Summary_sheet.xlsx", overwrite = T)

#

library(readxl)
library(openxlsx) #devtools::install_github("awalker89/openxlsx")
library(tidyverse)

# Create new workbook
#new_rass <- createWorkbook()


# New pollutant list
copc <- read_csv("updated_pollutant_list.csv")


# Pollutant groups
pbts <- read_csv("https://raw.githubusercontent.com/MPCA-air/health-values/master/pollutant_groups_PBTs.csv") %>%
        clean_names() %>%
        rename(CAS = cas)


copc <- left_join(copc, select(pbts, -pollutant_groups))



# Multi-media chemicals
multi <- read_csv("https://raw.githubusercontent.com/MPCA-air/health-values/master/multi_pathway_risk_factors.csv") %>%
         clean_names() %>%
         rename(CAS = cas) %>%
         rowwise() %>%
         mutate(is_multi = ifelse(sum(c(resident_noncancer, resident_cancer), na.rm = T) > 0, "X", ""))


copc <- copc %>% left_join(select(multi, CAS, is_multi))



# Early life adjustment factors
## 1.6x
early <- read_csv("https://raw.githubusercontent.com/MPCA-air/health-values/master/early_life_adjustment_ADAFs.csv") %>%
         clean_names() %>%
         rename(CAS = cas) %>%
         filter(tolower(additional_adaf_analysis_needed) == "yes")


#------------------------------#
# Create Risk SHEET
#------------------------------#

# Add a sheet
addWorksheet(new_rass, "Risk Calcs")

# Create content
new <- tibble()

new[1:3, 1] <- c("=Summary!A1", "No inputs on this page", " ")

new[, 2:22] <- ""

new[4, 1:22] <- c("CAS# or MPCA#", "Chemical Name",
                "Inhalation Hazard Quotients and Cancer Risks", rep("", 3),
                "Chronic Non-inhalation Pathway Hazard Quotients and Cancer Risks", rep("", 5),
                "Chronic Total Hazard Quotients and Cancer Risks (Inhalation + Non-inhalation)", rep("", 5),
                'Marked with "X" if applies', rep("", 3))

new[5, 1:22] <- c("", "",
                  "Acute", "Subchronic Noncancer", "Chronic Noncancer",	"Cancer",
                  rep(c("Farmer Noncancer",	"Farmer Cancer",	"Urban Gardener Noncancer",
                  "Urban Gardener Cancer", "Resident Noncancer", "Resident Cancer"), 2),
                  "Multimedia Chemical", "Persistent Bioaccumulative Toxicant",
                  "Respiratory Sensitizer", "Developmental Toxicant")

# Add total row
new[6, 1:2] <- c("", "Total")


# Add pollutant rows
for(i in 1:nrow(copc)) {
  new[6+i, 1]  <- copc[i, "CAS"]
  new[6+i, 2]  <- copc[i, "Chemical Name"]
  new[6+i, 19] <- copc[i, "is_multi"]
  new[6+i, 20] <- copc[i, "persistent_bioaccumulative_toxicants"]
  new[6+i, 21] <- copc[i, "respiratory_sensitizers"]
  new[6+i, 22] <- copc[i, "developmental_toxicants"]
}

# Add text
writeData(new_rass, "Risk Calcs", x = new, rowNames = F, colNames = F)


# Add version & formulas
writeFormula(new_rass,
             "Risk Calcs", x = "=Summary!A1", startRow = 1, startCol = 1)

# Total row formulas
for(column in 3:18) {

  letter <- LETTERS[column]

  print(letter)

  writeFormula(new_rass,
               "Risk Calcs",
               x        = gsub("H", letter,
                               paste0('=IF(AND(SUM(H7:H', 6+nrow(copc)+14, ')>0,$C$3=2869),SUM(H7:H', 6+nrow(copc)+14, '),',
                                      'IF(SUM(H7:H', 6+nrow(copc), ')>0,SUM(H7:H', 6+nrow(copc), '),""))')),
               startRow = 6,
               startCol = column)
}


# Pollutant row formulas

# Inhalation totals
for(i in 1:nrow(copc)) {

    # Get pollutant concentration row
    pollutant <- i + 5

    # Check for age adjustment factor
    age_factor <- max(c(1, 1.6*(copc[i, ]$CAS %in% early$CAS)), na.rm = T)

    # Acute
    writeFormula(new_rass,
                 "Risk Calcs",
                 x        = paste0("=IF(AND(Concs!C", pollutant, ">0,ISNUMBER('Tox Values'!G", 20+i, ")),Concs!C", pollutant, "/'Tox Values'!G", 20+i, ",\"\")"),
                 startRow = 6+i,
                 startCol = 3)

    # Subchronic
    writeFormula(new_rass,
                 "Risk Calcs",
                 x        = paste0("=IF(AND(Concs!F", pollutant, ">0,ISNUMBER('Tox Values'!AA", 20+i, ")),Concs!F", pollutant, "/'Tox Values'!AA", 20+i, ",\"\")"),
                 startRow = 6+i,
                 startCol = 4)

    # Chronic Hazard
    writeFormula(new_rass,
                 "Risk Calcs",
                 x        = paste0("=IF(AND(Concs!G", pollutant, ">0,ISNUMBER('Tox Values'!T", 20+i, ")),Concs!G", pollutant, "/'Tox Values'!T", 20+i, ",\"\")"),
                 startRow = 6+i,
                 startCol = 5)


    # Chronic Cancer
    ## Include Early Life Adjustment HERE????
    writeFormula(new_rass,
                 "Risk Calcs",
                 x        = paste0("=IF(AND(Concs!G", pollutant, ">0,ISNUMBER('Tox Values'!N", 20+i, ")),", age_factor, "*Concs!G", pollutant, "/'Tox Values'!N", 20+i, ",\"\")"),
                 startRow = 6+i,
                 startCol = 6)

}


# Non-Inhalation totals
for(i in 1:nrow(copc)) {

  # Get pollutant MPS
  mps_pollutant <- i + 19

  # Get pollutant risk
  rsk_pollutant <- i + 6

  # Farmer
  ## Non-Cancer
  writeFormula(new_rass,
               "Risk Calcs",
               x        = paste0("=IF(AND('MPS Factors'!G", mps_pollutant,
                                 ">0,ISNUMBER('Risk Calcs'!E", rsk_pollutant,
                                 ")),'MPS Factors'!G", mps_pollutant, "*'Risk Calcs'!E", rsk_pollutant, ",\"\")"),
               startRow = 6+i,
               startCol = 7)

  ## Cancer
  writeFormula(new_rass,
               "Risk Calcs",
               x        = paste0("=IF(AND('MPS Factors'!H", mps_pollutant,
                                 ">0,ISNUMBER('Risk Calcs'!F", rsk_pollutant,
                                 ")),'MPS Factors'!H", mps_pollutant, "*'Risk Calcs'!F", rsk_pollutant, ",\"\")"),
               startRow = 6+i,
               startCol = 8)


  # Urban Gardener
  ## Non-Cancer
  writeFormula(new_rass,
               "Risk Calcs",
               x        = paste0("=IF(AND('MPS Factors'!E", mps_pollutant,
                                 ">0,ISNUMBER('Risk Calcs'!E", rsk_pollutant,
                                 ")),'MPS Factors'!E", mps_pollutant, "*'Risk Calcs'!E", rsk_pollutant, ",\"\")"),
               startRow = 6+i,
               startCol = 9)

  ## Cancer
  writeFormula(new_rass,
               "Risk Calcs",
               x        = paste0("=IF(AND('MPS Factors'!F", mps_pollutant,
                                 ">0,ISNUMBER('Risk Calcs'!F", rsk_pollutant,
                                 ")),'MPS Factors'!F", mps_pollutant, "*'Risk Calcs'!F", rsk_pollutant, ",\"\")"),
               startRow = 6+i,
               startCol = 10)



  # Resident
  ## Non-Cancer
  writeFormula(new_rass,
               "Risk Calcs",
               x        = paste0("=IF(AND('MPS Factors'!C", mps_pollutant,
                                 ">0,ISNUMBER('Risk Calcs'!E", rsk_pollutant,
                                 ")),'MPS Factors'!C", mps_pollutant, "*'Risk Calcs'!E", rsk_pollutant, ",\"\")"),
               startRow = 6+i,
               startCol = 11)

  ## Cancer
  writeFormula(new_rass,
               "Risk Calcs",
               x        = paste0("=IF(AND('MPS Factors'!D", mps_pollutant,
                                 ">0,ISNUMBER('Risk Calcs'!F", rsk_pollutant,
                                 ")),'MPS Factors'!D", mps_pollutant, "*'Risk Calcs'!F", rsk_pollutant, ",\"\")"),
               startRow = 6+i,
               startCol = 12)

}




# All media risk totals
for(i in 1:nrow(copc)) {

  # Get pollutant risk
  rsk_pollutant <- 6+i

  # Farmer
  ## Non-Cancer
  writeFormula(new_rass, "Risk Calcs",
               x        = paste0("=E", mps_pollutant,
                                 "+G", rsk_pollutant),
               startRow = 6+i,
               startCol = 13)

  ## Cancer
  writeFormula(new_rass,
               "Risk Calcs",
               x        = paste0("=F", mps_pollutant,
                                 "+H", rsk_pollutant),
               startRow = 6+i,
               startCol = 14)


  # Urban Gardener
  ## Non-Cancer
  writeFormula(new_rass,
               "Risk Calcs",
               x        = paste0("=E", mps_pollutant,
                                 "+I", rsk_pollutant),
               startRow = 6+i,
               startCol = 15)

  ## Cancer
  writeFormula(new_rass,
               "Risk Calcs",
               x        = paste0("=F", mps_pollutant,
                                 "+J", rsk_pollutant),
               startRow = 6+i,
               startCol = 16)



  # Resident
  ## Non-Cancer
  writeFormula(new_rass,
               "Risk Calcs",
               x        = paste0("=E", mps_pollutant,
                                 "+K", rsk_pollutant),
               startRow = 6+i,
               startCol = 17)

  ## Cancer
  writeFormula(new_rass,
               "Risk Calcs",
               x        = paste0("=F", mps_pollutant,
                                 "+L", rsk_pollutant),
               startRow = 6+i,
               startCol = 18)

}


# Save
#saveWorkbook(new_rass, file = "03_Risk_Calcs_sheet.xlsx", overwrite = T)

#

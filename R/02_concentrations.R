library(readxl)
library(openxlsx) #devtools::install_github("awalker89/openxlsx")
library(tidyverse)

# Create new workbook
#new_rass <- createWorkbook()


# Load default RASS template
#emit <- readWorkbook("Test_R_RASS-Copy.xlsx", sheet = 3)

# New pollutant list
copc <- read_csv("updated_pollutant_list.csv")


#------------------------------#
# Create Concs SHEET
#------------------------------#

# Add a sheet
addWorksheet(new_rass, "Concs")

# Create content
new <- tibble()

new[1:3, 1] <- c("=Summary!A1", "No inputs needed on this page", " ")

new[, 2:132] <- ""

new[4, 1:(2+26*5)] <- c("Air Concentrations in ug/m3", "",
                        rbind(c("Total - All stacks",
                                paste0("Stack #", 1:25)),
                              c(""), c(""), c(""), c("")) %>%
                                as.vector())

new[5, 1:(2+26*5)] <- c("CAS# or MPCA#", "Chemical Name", rep(c("C (1-hr)", "C (3-hr))", "C (24-hr)", "C (monthly))", "C (annual)"), 26))

# Add pollutant rows
for(i in 1:nrow(copc)) {
        new[5+i, 1] <- copc[i, "CAS"]
        new[5+i, 2] <- copc[i, "Chemical Name"]
}

new_conc <- new



# Add text
writeData(new_rass, "Concs", x = new_conc, rowNames = F, colNames = F)

# Add version & formulas
writeFormula(new_rass,
             "Concs", x = "=Summary!A1", startRow = 1, startCol = 1)


# Stack Totals
for(i in 1:nrow(copc)) {

        print(copc[i, ]$"Chemical Name")

        # 1-HR
        writeFormula(new_rass,
                     "Concs",
                     x        = gsub("14", 5+i, "=H14+M14+R14+W14+AB14+AG14+AL14+AQ14+AV14+BA14+BF14+BK14+BP14+BU14+BZ14+CE14+CJ14+CO14+CT14+CY14+DD14+DI14+DN14+DS14+DX14"),
                     startRow = 5+i,
                     startCol = 3)

        # 3-HR
        writeFormula(new_rass,
                     "Concs",
                     x        = gsub("14", 5+i, "=I14+N14+S14+X14+AC14+AH14+AM14+AR14+AW14+BB14+BG14+BL14+BQ14+BV14+CA14+CF14+CK14+CP14+CU14+CZ14+DE14+DJ14+DO14+DT14+DY14"),
                     startRow = 5+i,
                     startCol = 4)

        # 24-HR
        writeFormula(new_rass,
                     "Concs",
                     x        = gsub("14", 5+i, "=J14+O14+T14+Y14+AD14+AI14+AN14+AS14+AX14+BC14+BH14+BM14+BR14+BW14+CB14+CG14+CL14+CQ14+CV14+DA14+DF14+DK14+DP14+DU14+DZ14"),
                     startRow = 5+i,
                     startCol = 5)

        # Monthly
        writeFormula(new_rass,
                     "Concs",
                     x        = gsub("14", 5+i, "=K14+P14+U14+Z14+AE14+AJ14+AO14+AT14+AY14+BD14+BI14+BN14+BS14+BX14+CC14+CH14+CM14+CR14+CW14+DB14+DG14+DL14+DQ14+DV14+EA14"),
                     startRow = 5+i,
                     startCol = 6)

        # Annual
        writeFormula(new_rass,
                     "Concs",
                     x        = gsub("14", 5+i, "=L14+Q14+V14+AA14+AF14+AK14+AP14+AU14+AZ14+BE14+BJ14+BO14+BT14+BY14+CD14+CI14+CN14+CS14+CX14+DC14+DH14+DM14+DR14+DW14+EB14"),
                     startRow = 5+i,
                     startCol = 7)

}


# Stack specific
for(i in 1:nrow(copc)) {

   print(copc[i, ]$"Chemical Name")

   for(stack in 1:25) {

        # Get pollutant emissions row
        pollutant <- i + 10

        # Stack emissions column
        stack_emit <- c(LETTERS[-c(1:4)], paste0("A", LETTERS), "BA", "BB")[c((stack*2-1), (stack*2))]

        # Dispersion column
        stack_disp <- c(LETTERS[-c(1:2)], "AA")[stack]


        # 1-HR
        writeFormula(new_rass,
                     "Concs",
                     x        = paste0("=IF(AND(ISNUMBER('Emissions (start here)'!$", stack_emit[1], pollutant, "),",
                                       "ISNUMBER('Dispersion'!$", stack_disp, "$14)),",
                                       "'Emissions (start here)'!$", stack_emit[1], pollutant,
                                       "*453.59/3600*'Dispersion'!$", stack_disp, "$14,0)"),
                     startRow = 5+i,
                     startCol = 3+5*stack)

        # 3-HR
        writeFormula(new_rass,
                     "Concs",
                     x        = paste0("=IF(AND(ISNUMBER('Emissions (start here)'!$", stack_emit[1], pollutant, "),",
                                       "ISNUMBER('Dispersion'!$", stack_disp, "$15)),",
                                       "'Emissions (start here)'!$", stack_emit[1], pollutant,
                                       "*453.59/3600*'Dispersion'!$", stack_disp, "$15,0)"),
                     startRow = 5+i,
                     startCol = 4+5*stack)

        # 24-HR
        writeFormula(new_rass,
                     "Concs",
                     x        = paste0("=IF(AND(ISNUMBER('Emissions (start here)'!$", stack_emit[1], pollutant, "),",
                                       "ISNUMBER('Dispersion'!$", stack_disp, "$16)),",
                                       "'Emissions (start here)'!$", stack_emit[1], pollutant,
                                       "*453.59/3600*'Dispersion'!$", stack_disp, "$16,0)"),
                     startRow = 5+i,
                     startCol = 5+5*stack)

        # Monthly
        writeFormula(new_rass,
                     "Concs",
                     x        = paste0("=IF(AND(ISNUMBER('Emissions (start here)'!$", stack_emit[2], pollutant, "),",
                                       "ISNUMBER('Dispersion'!$", stack_disp, "$17)),",
                                       "'Emissions (start here)'!$", stack_emit[2], pollutant,
                                       "*2000*453.59/8760/3600*'Dispersion'!$", stack_disp, "$17,0)"),
                     startRow = 5+i,
                     startCol = 6+5*stack)

        # Annual
        writeFormula(new_rass,
                     "Concs",
                     x        = paste0("=IF(AND(ISNUMBER('Emissions (start here)'!$", stack_emit[2], pollutant, "),",
                                       "ISNUMBER('Dispersion'!$", stack_disp, "$18)),",
                                       "'Emissions (start here)'!$", stack_emit[2], pollutant,
                                       "*2000*453.59/8760/3600*'Dispersion'!$", stack_disp, "$18,0)"),
                     startRow = 5+i,
                     startCol = 7+5*stack)

}}


# Save
#saveWorkbook(new_rass, file = "02_Concs_sheet.xlsx", overwrite = T)

#


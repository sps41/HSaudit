library("tidyverse")
library(openxlsx)

rm(list = ls())
#setwd("~/Steve/work/Computational projects/psoriasis_BAD_audit")
dir <- paste(getwd(),"/Trust responses",sep="")
files <- list.files(dir,full.names = T)
###Read first file and create data frames
###Read in patient data: hospital data in cols C and D, then drop the blank rows
Pat1a <- read_excel(files[1],sheet = "Option 1",range = "C2:D17",col_names = T)
Pat1a <- Pat1a[c(3,6,9,12,15),]
###Read in patient data: patient record data in cols H to AE, then drop the blank rows
Pat2a <- read_excel(files[1],sheet = "Option 1",range = "H4:AE17",col_names = T)
Pat2a <- Pat2a[c(1,4,7,10,13),]
##Join them up
patient_data <- cbind(Pat1a,Pat2a)

###Read in clinician data
Clin1 <- read_excel(files[1],sheet = "Option 1",range = "C20:M21",col_names = T)
clinician_data <- Clin1[,c(1,5,9,11)]


####Read in data for all others
for(f in 2:length(files)){
  print(paste("Starting:",files[f]))
  
  ###Read in patient data: hospital data in cols C and D, then drop the blank rows
  tmp1a <- read_excel(files[f],sheet = "Option 1",range = "C3:D17",col_names = F)
  tmp1a <- tmp1a[c(3,6,9,12,15),]
  ###Read in patient data: patient record data in cols H to AE, then drop the blank rows
  tmp2a <- read_excel(files[f],sheet = "Option 1",range = "H5:AE17",col_names = F)
  tmp2a <- tmp2a[c(1,4,7,10,13),]
  ##Join them up
  tmpAll <- cbind(tmp1a,tmp2a)
  colnames(tmpAll) <- colnames(patient_data)
  patient_data <- rbind(patient_data,tmpAll)

    ###Read in clinician data
  Clin1a <- read_excel(files[f],sheet = "Option 1",range = "C21:M21",col_names = F)
  Clin1a <- Clin1a[,c(1,5,9,11)]
  colnames(Clin1a) <- colnames(clinician_data)
  clinician_data <- rbind(clinician_data,Clin1a)
  print(paste("Done:",files[f]))
}

Unified_data <- patient_data
expanded_clinician <- rep(clinician_data[[3]],each=5)
Unified_data$contact <- expanded_clinician

write.xlsx(list(Unified_data,clinician_data,patient_data), file = "Combined_HS_audit_data.xlsx", borders = "columns")

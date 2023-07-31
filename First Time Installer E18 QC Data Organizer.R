#Install R 3.4.4 and Rtools before running this script
#First time installation of QC Data Organize Script
install.packages("purrr")
install.packages("stringr")
install.packages("plyr")
install.packages("gdata")
install.packages("dplyr")
install.packages("lubridate")
install.packages("miscTools")
install.packages("DataCombine")
install.packages("splitstackshape")
install.packages("openxlsx")
install.packages("ggplot2")

#Create working directory folder, this will also be checked by the organizer script.

wd <- "E18 QC Data Organizer with R"

if (!basename(getwd()) == wd){
  dir.create(paste(getwd(), "/Documents/Emerald 18 QC Data Organize with R", sep = ""))
  
  setwd(paste(getwd(), "/Documents/Emerald 18 QC Data Organize with R", sep = "")) 
  }


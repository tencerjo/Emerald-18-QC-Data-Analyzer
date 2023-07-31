#This script creates Excel files, it will NOT overwrite previously created files with the same name, so move or delete the old files it created.

#TO RUN THIS SCRIPT IN RSTUDIO CLICK IN THE CODE AREA AND PRESS (CTRL + ALT + R) OR go to Code -> Run Region -> Run All

#TO CHANGE THE PARAMETERS OF THE OUTPUT FILE, ADD DESIRED PARAMETERS TO finalcols FROM THE LIST OF AVAILABLE PARAMETERS FURTHER DOWN.
#DO NOT EDIT BELOW THE #HASH# LINE UNLESS YOU KNOW WHAT YOU ARE DOING

#Clear the workspace to remove any artifacts from last session.
rm(list = ls())

#IF THE SCRIPT CANNOT FIND YOUR ASSAY/QC FILES USE THIS TO SET THE FOLDER THAT CONTAINS ALL QC/ASSAY FILES IN 
# IT OR IN A SUBFOLDER.
#COPY AND PASTE THE FULL FOLDER LOCATION AND CHANGE "\" TO "/" AND SURROUND PATH WITH QUOTES
#EXAMPLE: C:\Users\[USERNAME]\Documents\Emerald 18 QC Data Organize with R

#setwd("C:/Users/[USERNAME]/Documents/Emerald 18 QC Data Organize with R")

#set to FALSE to turn excel graph creation off, set to TRUE to turn graphs on (very slow)
graphs <- FALSE

#Start row for assay values in the created excel file (buffer for comments at top)
assaystart <- 6

#buffer of empty cells between QC values and assay values
datastart <- 13

#Final rearrangement of columns QC. use "RBC_LL", "MCV_NL" for RBC low level values and MCV normal level values etc.

# Available parameters to choose from for finalcols:
# "SN"        "Level"     "Lot"       "WBC_LL"    "GRAN_LL"   "%G_LL"     "LYM_LL"    "%L_LL"     "MID_LL"    "%M_LL"     "RBC_LL"    "HGB_LL"   
# "HCT_LL"    "MCV_LL"    "MCH_LL"    "MCHC_LL"   "RDW_LL"    "PLT_LL"    "MPV_LL"    "PCT_LL"    "PDW_LL"    "Date"      "Time_LL"   "Comments"  
# "ShortDate" "SN"        "Level"     "Lot"       "WBC_NL"    "GRAN_NL"   "%G_NL"     "LYM_NL"    "%L_NL"     "MID_NL"    "%M_NL"     "RBC_NL"   
# "HGB_NL"    "HCT_NL"    "MCV_NL"    "MCH_NL"    "MCHC_NL"   "RDW_NL"    "PLT_NL"    "MPV_NL"    "PCT_NL"    "PDW_NL"    "Date"      "Time_NL"  
# "Comments"   "ShortDate" "SN"        "Level"     "Lot"       "WBC_HL"    "GRAN_HL"   "%G_HL"     "LYM_HL"    "%L_HL"     "MID_HL"    "%M_HL"    
# "RBC_HL"    "HGB_HL"    "HCT_HL"    "MCV_HL"    "MCH_HL"    "MCHC_HL"   "RDW_HL"    "PLT_HL"    "MPV_HL"    "PCT_HL"    "PDW_HL"    "Date"     
# "Time_HL"   "Comments"   "ShortDate" "datetime_LL" "datetime_NL" "datetime_HL"

# Presets for less typing, use to define finalcols
# do not edit these, only edit finalcols object
# (do NOT put quotes around these when put in finalcols), objects are CaSe SeNsItIvE!
RBC <- c("RBC_LL", "RBC_NL", "RBC_HL")
MCV <- c("MCV_LL", "MCV_NL", "MCV_HL")
HGB <- c("HGB_LL", "HGB_NL", "HGB_HL")
HCT <- c("HCT_LL", "HCT_NL", "HCT_HL")
MCH <- c("MCH_LL", "MCH_NL", "MCH_HL")
MCHC <- c("MCHC_LL", "MCHC_NL", "MCHC_HL")
RDW <- c("RDW_LL", "RDW_NL", "RDW_HL")

PLT <- c("PLT_LL", "PLT_NL", "PLT_HL")
MPV <- c("MPV_LL", "MPV_NL", "MPV_HL")
PCT <- c("PCT_LL", "PCT_NL", "PCT_HL")
PDW <- c("PDW_LL", "PDW_NL", "PDW_HL")

WBC <- c("WBC_LL", "WBC_NL", "WBC_HL")
GRAN <- c("GRAN_LL", "GRAN_NL", "GRAN_HL", "%G_LL", "%G_NL", "%G_HL")
LYM <- c("LYM_LL", "LYM_NL", "LYM_HL", "%L_LL", "%L_NL", "%L_HL")
MID <- c("MID_LL", "MID_NL", "MID_HL", "%M_LL", "%M_NL", "%M_HL")


CBC <- c(WBC, RBC, HGB, HCT, MCV, MCH, MCHC, RDW, PLT, MPV, PCT, PDW)
DIFF <- c(WBC, GRAN, LYM, MID)
ALL <- c(DIFF, RBC, HGB, HCT, MCV, MCH, MCHC, RDW, PLT, MPV, PCT, PDW)
DATES <- c("datetime_LL", "datetime_NL", "datetime_HL")

#Define finalcols object for excel output
#To use presets, do NOT use quotes, but individual columns MUST have "quotes".

finalcols <- c("Date", RBC, MCV, PLT, "Comments", "ShortDate", DATES)



#MAKE A COPY OF THIS FILE BEFORE EDITING CODE
###DO NOT EDIT BELOW THIS LINE UNLESS YOU KNOW WHAT YOU ARE DOING ####################################################################################################################

#The assay value and conditional formatting in excel adjusts automaticaly! DO NOT CHANGE BELOW. THESE WILL ADJUST TO YOUR PICKS FOR FINALCOLS.

#Get finalcol names that do not include "Comments", "Time", "ShortDate", "SN", and "Level" columns,
#include lot number, requested columns, and expiration date ("Lot") column and set that to QC column names to get

avcols  <- c(finalcols[c(-grep("Comments|Time|SN|Level|ShortDate|datetime", finalcols))], "Lot")

#Only value rows for conditional formatting in Excel
valuecols  <- c(finalcols[c(-grep("Date|Comments|Time|SN|Level|ShortDate|datetime", finalcols))])


#Load necessary libraries
library("purrr")
library("plyr")
library("gdata")
library("dplyr")
library("stringr")
library("lubridate")
library("miscTools")
library("DataCombine")
library("splitstackshape")
library("openxlsx")
library("ggplot2")
#library("plotly")
#library("remotes")
#library("later")
#library("data.table")

#redefine q (quit) so R profile/history isn't saved
q <- function (save = "no", status = 0, runLast = TRUE)
  .Internal(quit(save, status, runLast))

#Time sorting time differential, skip row if greater than this value
#timediff <- (as.POSIXct("00:45:00", format = "%H:%M:%S") - as.POSIXct("00:00:00", format = "%H:%M:%S"))

#list all QC DAT files
qcfiles <- list.files(path = ".", pattern = "*[:alpha:]*[:alpha:]*[:alpha:]*[:alpha:]*QC.DAT$", recursive = TRUE)

#Stop if no QC files
if (length(qcfiles) == 0){stop("Place the AB18 QC data folder into the same folder or subfolder of the E18 QC Data Organizer.R file")}

#Stop if a qc file not in standard L/N/H### format
#if (word()) {
#  
#}


#Assay Value check inserted for faster Stop if no Assay Value files in folder.
#Read Control Assay Value Files
avfiles <- list.files(path = ".", pattern = "*.QC[[:alpha:]]", recursive = TRUE )

#Stop if no Assay Value files
if(length(avfiles) < 2){stop("Place the Assay Value files ####.QC in the same folder or subfolder as the E18 QC Data Organizer.R file")}



#just sn/sn/eqc/qc lot of path to ignore dir, str_extract because R's grep is garbage and just doesn't work..
snlotpath <- str_extract(qcfiles, pattern = "[:digit:][:digit:][:digit:][:digit:][:digit:][:digit:]/EQC/[LNH]?[:digit:][:digit:][:digit:][:alnum:]?[:alnum:]?[:alnum:]?[:alnum:]?/[:alpha:]?[:alpha:]?[:alpha:]?[:alpha:]?[:alpha:]+[:alpha:]+QC.DAT$")

#get file info for date modified
finfo <- file.info(qcfiles, extra_cols = TRUE)

#duplicate file path name and keep date modified columns for sorting
files <- cbind(as.character(qcfiles), as.character(snlotpath), data.frame(finfo$mtime))

#remove rows with NAs, then sort by name and date modified, keep only unique
NAfiles<- na.omit(files)

#sort by date, decreasing = TRUE so oldest duplicates are deleted instead of newest
orfiles <- NAfiles[order(NAfiles[,3], decreasing = TRUE),]

#remove duplicates, keeping most recent
ufiles <- orfiles[!duplicated(orfiles[["as.character(snlotpath)"]]),]

#fix class AGAIN to be character not factor
ufiles[,1] <- as.character(ufiles[,1])
ufiles[,2] <- as.character(ufiles[,2])

#create an empty list to use in for function to get 
list.qc <- list()

#sort the files into a list, read the tables keeping only data, date, and time.

for (i in 1:length(ufiles[,1])) {
  
  print(paste("If \"no lines avaliable in input error\" please check or delete the last file name printed here.
                    File may be empty and unreadable, please check and delete file", ufiles[i,1]))
  
  list.qc[[i]] <- read.table(file = ufiles[i,1], header = FALSE, sep =  "\t", skip = 3, na.strings = c("--","---", "----", "-----"),
                             colClasses = c("NULL", "numeric","NULL", "numeric","NULL", "numeric","NULL", "numeric","NULL", "numeric",
                                            "NULL", "numeric","NULL", "numeric","NULL", "numeric","NULL", "numeric","NULL", "numeric",
                                            "NULL", "numeric","NULL", "numeric","NULL", "numeric","NULL", "numeric","NULL", "numeric","NULL",
                                            "numeric","NULL", "numeric","NULL", "numeric", NA, NA)
                             , skipNul = TRUE)
}


#Faster read in using lapply, but no error handling
# list.qc <- lapply(ufiles[,1], read.table, header = FALSE, sep =  "\t", skip = 3, na.strings = c("--","---", "----", "-----"),
#                   colClasses = c("NULL", "numeric","NULL", "numeric","NULL", "numeric","NULL", "numeric","NULL", "numeric",
#                                  "NULL", "numeric","NULL", "numeric","NULL", "numeric","NULL", "numeric","NULL", "numeric",
#                                  "NULL", "numeric","NULL", "numeric","NULL", "numeric","NULL", "numeric","NULL", "numeric","NULL",
#                                  "numeric","NULL", "numeric","NULL", "numeric", NA, NA)
#                   , skipNul = TRUE)


#use purr map here to set date format
#Not using lubridate because lubridate cannot coerce NA to class date, use lubridate after sorting

for (i in 1:length(ufiles[,1])){
  #set date column to Date class for comparison later.
  list.qc[[i]][,19] <- format(as.Date(list.qc[[i]][,19], format = "%m/%d/%y",origin = "00/00/00"), format = "%m/%d/%y")
  
}

#CHECK VALUE OF i IF ERROR MESSAGE, AND CHECK ufiles[i] TO SEE MESSED UP TABLE.

Comments <- NA

#Extract Serial Number
sn <- word(ufiles[,2],-4, sep = fixed("/"))

#Extract lot and level
level <- word(ufiles[,2], -2, sep = fixed("/"))

#Remove front letter of lot
ulevel <- sub("^[[:alpha:]]", "",level)

#Keep only unique and sort in order, for continuous wb creation later
ulevel <- sort(unique(ulevel))

#sort and keep only unique Serial Numbers, sorted for later continuous workbook creation.
usn <- sort(unique(sn))

#split L/N/H and lot # into 2 variables for labeling tables
lot <- sub("^[[:alpha:]]", "",level)
level <- str_extract(level, "^[[:alpha:]]")

#join names for snlevellot column to be split later
snlevellot <- paste(sn, level, lot, sep = "_")

#Label tables with names
names(list.qc) <- snlevellot


#define column names for renaming
#                 1       2     3       4     5     6     7     8       9     10      11      12    13      14    
QCheaders <- c( "WBC", "GRAN", "%G", "LYM", "%L", "MID", "%M", "RBC", "HGB", "HCT", "MCV", "MCH", "MCHC", "RDW",
                # 15      16    17      18    19      20        21
                "PLT", "MPV", "PCT", "PDW", "Date", "Time", "datetime")


#add datetime column
#name headers to the list
for (i in 1:length(list.qc))
{
  #just placeholder to add LL, NL, HL column names to datetime
  list.qc[[i]][["datetime"]] <- NA
  
  names(list.qc[[i]]) <- QCheaders
}


#Add Low, Normal, High to column names
for (i in 1:length(list.qc)) {

  #paste on to column names "LL", "NL", "HL". (Low Level, Normal Level, High Level)
  #Keep "Date" column named "Date" for sorting

  #Low
  if (grepl("*L", names(list.qc[i])))
  {names(list.qc[[i]]) <- sub("Date_LL", "Date", paste(names(list.qc[[i]]), "LL", sep = "_"))}

  #Normal
  if (grepl("*N", names(list.qc[i])))
  {names(list.qc[[i]]) <- sub("Date_NL", "Date", paste(names(list.qc[[i]]), "NL", sep = "_"))}

  #High
  if (grepl("*H", names(list.qc[i])))
  {names(list.qc[[i]]) <- sub("Date_HL", "Date", paste(names(list.qc[[i]]), "HL", sep = "_"))}
  
  }


#Add column with sn, lot and level to first column in table
#kept to preserve column numbers, delete during renaming and adding comment field.
for (i in 1:length(list.qc)) {
  list.qc[[i]] <- cbind(snlevellot[i], list.qc[[i]], Comments)
  
  list.qc[[i]][["snlevellot[i]"]] <- as.character(list.qc[[i]][["snlevellot[i]"]])
}


#search for tables for each Serial Number
#for (i in 1:) {
#  contains(match = i, vars = list,qc)
#}

#sort by L, N, H, then number #THIS ONLY KEEPS LOW NORMAL HIGH, NO SPECIAL LOTS
#L <- grep("^L", ulevel, value = TRUE)
#N <- grep("^N", ulevel, value = TRUE)
#H <- grep("^H", ulevel, value = TRUE)

#instead of using L, N, H just keep unique lot numbers, remove first letter and keep unique.

#keep only lots in L/N/H### format
#level <- grep("^[[LNH]][[:digit:]][[:digit:]][[:digit:]][[:digit:]]", level, value = TRUE)

#Compare date columns and insert row when date is > in other but if NA next

#Create duplicate list.qc to modify dates on
firstdatesort <- list.qc

#NA value vector same length as # of columns in table  list/vector for InsertRow
NArow <- 1:length(names(firstdatesort[[1]])) * NA

#Final changes to list.qc before publishing

#serial number counter
for (j in 1:length(usn)){
  
  #QClevel counter
  for (i in 1:length(ulevel)){
    
    #assign search "level sn" change lot to L N H for each element
    L <- paste(usn[j], "L", ulevel[i], sep = "_")
    N <- paste(usn[j], "N", ulevel[i], sep = "_")
    H <- paste(usn[j], "H", ulevel[i], sep = "_")
    
    #check if sn/qc combo exists in list.qc
    if (!exists(L, where = firstdatesort)) {next}
    if (!exists(N, where = firstdatesort)) {next}
    if (!exists(H, where = firstdatesort)) {next}
    
    #compare date rows to set index to longest row
    #unable to change iteration number inside loop, set it to arbitrarily high number max datecount *3
     datecountL <- length(firstdatesort[[L]][["Date"]])
     datecountN <- length(firstdatesort[[N]][["Date"]])
     datecountH <- length(firstdatesort[[H]][["Date"]])
    
    
    for (k in 1:max(datecountL*3, datecountN*3, datecountH*3)) {
      
      #Low to Normal check
      #new table edited
      if (if_else(firstdatesort[[L]][k,"Date"] > firstdatesort[[N]][k,"Date"], TRUE, FALSE, missing = FALSE)){firstdatesort[[L]] <- InsertRow(firstdatesort[[L]], NArow, RowNum = k)
      }
      
      #Low to High check                                                   
      #new table edited
      if (if_else(firstdatesort[[L]][k,"Date"] > firstdatesort[[H]][k,"Date"], TRUE, FALSE, missing = FALSE)){firstdatesort[[L]] <- InsertRow(firstdatesort[[L]], NArow, RowNum = k)
      }
      
      #Normal to Low check                                                  
      #new table edited
      if (if_else(firstdatesort[[N]][k,"Date"] > firstdatesort[[L]][k,"Date"], TRUE, FALSE, missing = FALSE)){firstdatesort[[N]] <- InsertRow(firstdatesort[[N]], NArow, RowNum = k)
      }
      
      #Normal to High check  
      #new table edited
      if (if_else(firstdatesort[[N]][k,"Date"] > firstdatesort[[H]][k,"Date"], TRUE, FALSE, missing = FALSE)){firstdatesort[[N]] <- InsertRow(firstdatesort[[N]], NArow, RowNum = k)
      }
      
      #High to Low Check
      #new table edited
      if (if_else(firstdatesort[[H]][k,"Date"] > firstdatesort[[L]][k,"Date"], TRUE, FALSE, missing = FALSE)){firstdatesort[[H]] <- InsertRow(firstdatesort[[H]], NArow, RowNum = k)
      }
      
      #High to Normal Check
      #new table edited
      if (if_else(firstdatesort[[H]][k,"Date"] > firstdatesort[[N]][k,"Date"], TRUE, FALSE, missing = FALSE)){firstdatesort[[H]] <- InsertRow(firstdatesort[[H]], NArow, RowNum = k)
      }
      
    # ###Less than check ### Moves the other file down that is not equal but not less than (so must be greater than)
    # ##This code not needed after final datecount length check and coalesce
    #  
    #   #Normal to Low check
    #   #new table edited
    #   if (if_else(firstdatesort[[N]][k,"Date"] < firstdatesort[[L]][k,"Date"], TRUE, FALSE, missing = FALSE)){firstdatesort[[L]] <- InsertRow(firstdatesort[[L]], NArow, RowNum = k)
    #    k=k-1}
    #   
    #   #High to Low check                                                  
    #   #new table edited
    #   if (if_else(firstdatesort[[H]][k,"Date"] < firstdatesort[[L]][k,"Date"], TRUE, FALSE, missing = FALSE)){firstdatesort[[L]] <- InsertRow(firstdatesort[[N]], NArow, RowNum = k)
    #    k=k-1}
    #   
    #   #Low to Normal check                                                   
    #   #new table edited
    #   if (if_else(firstdatesort[[L]][k,"Date"] < firstdatesort[[N]][k,"Date"], TRUE, FALSE, missing = FALSE)){firstdatesort[[N]] <- InsertRow(firstdatesort[[L]], NArow, RowNum = k)
    #    k=k-1}
    #   
    #   #High to Normal check  
    #   #new table edited
    #   if (if_else(firstdatesort[[H]][k,"Date"] < firstdatesort[[N]][k,"Date"], TRUE, FALSE, missing = FALSE)){firstdatesort[[N]] <- InsertRow(firstdatesort[[N]], NArow, RowNum = k)
    #    k=k-1}
    #   
    #   #High to Low Check
    #   #new table edited
    #   if (if_else(firstdatesort[[L]][k,"Date"] < firstdatesort[[H]][k,"Date"], TRUE, FALSE, missing = FALSE)){firstdatesort[[H]] <- InsertRow(firstdatesort[[H]], NArow, RowNum = k)
    #    k=k-1}
    #   
    #   #High to Normal Check
    #   #new table edited
    #   if (if_else(firstdatesort[[N]][k,"Date"] < firstdatesort[[H]][k,"Date"], TRUE, FALSE, missing = FALSE)){firstdatesort[[H]] <- InsertRow(firstdatesort[[H]], NArow, RowNum = k)
    #    k=k-1}
      
      
          }
    
    #j sn/level loop after sorting - Make tables same length (add NA if necessary), then merge dates for L N H here  each lot/sn after date sort
    
     #recount table length
     repeat{
      datecountL <- length(firstdatesort[[L]][["Date"]])
      datecountN <- length(firstdatesort[[N]][["Date"]])
      datecountH <- length(firstdatesort[[H]][["Date"]])
      
     #Add rows to bottom of table until table length is equal
      #low check
      if (datecountL < max(datecountL, datecountN, datecountH)){
        
        firstdatesort[[L]] <- rbind(firstdatesort[[L]], NArow)
      }
      
      #normal check
      if (datecountN < max(datecountL, datecountN, datecountH)){
        
        firstdatesort[[N]] <- rbind(firstdatesort[[N]], NArow)
      }
      
      #high check
      if (datecountH < max(datecountL, datecountN, datecountH)){
        
        firstdatesort[[H]] <- rbind(firstdatesort[[H]], NArow)
      }
        
        #break when all column lengths are equal
        if(datecountL == datecountN & datecountN == datecountH){break}
  }
    
    ##if (if_else(firstdatesort[[L]][["Date"]] == firstdatesort[[N]][["Date"]] & firstdatesort[[H]][["Date]] == firstdatesort[[L]][["Date"]], TRUE, FALSE, missing = TRUE)){

       #unlist to vectorize, retains Date type but no longer a dataframe
       Lunlist <- unlist(firstdatesort[[L]][["Date"]])
       Nunlist <- unlist(firstdatesort[[N]][["Date"]])
       Hunlist <- unlist(firstdatesort[[H]][["Date"]])
       
       #coalesce to merge dates, only if vectors lengths are equal
       #assign new name and coerce to dataframe
       #assign(paste(paste("Sdate", ulevel[i], sep = ""), usn[j], sep = "_"), as.data.frame(coalesce(Lunlist, Nunlist, Hunlist)))
       
       #replace old dates or set equal to new, [["Date]]
       #may need to change back to as.data.frame(coalesce())
       firstdatesort[[L]][["Date"]] <- coalesce(Lunlist, Nunlist, Hunlist)
       firstdatesort[[N]][["Date"]] <- coalesce(Lunlist, Nunlist, Hunlist)
       firstdatesort[[H]][["Date"]] <- coalesce(Lunlist, Nunlist, Hunlist)
       
       #} for if statement if used, tables are already checked for length
  }
} 


# 
# #Compare Date AND Time columns and insert row when date is = AND time is > ~40? in other but skip if NA 
# 
# #Create duplicate list from already modified date list
# firsttimesort <- firstdatesort
# 
# 
# #set time class for time column
# #change this to lapply in the future
# for (i in 1:length(firsttimesort)) {
# 
#   firsttimesort[[i]][,21] <- as.POSIXct(as.character(firsttimesort[[i]][,21]), format = "%H:%M:%S")
# }
# 
# 
# for (j in 1:length(usn)){
# 
#   for (i in 1:length(ulevel)){
# 
#     #assign search "level sn" change lot to L N H for each element
#     L <- paste(usn[j], "L", ulevel[i], sep = "_")
#     N <- paste(usn[j], "N", ulevel[i], sep = "_")
#     H <- paste(usn[j], "H", ulevel[i], sep = "_")
# 
#     ##check if sn/qc combo exists in list.qc
# if (!exists(L, where = firsttimesort)) {next}
# if (!exists(N, where = firsttimesort)) {next}
# if (!exists(H, where = firsttimesort)) {next}
# 
#     #compare time rows to set index to longest row
#     #Time = col 21
#     timecountL <- length(firsttimesort[[L]][,21])
#     timecountN <- length(firsttimesort[[N]][,21])
#     timecountH <- length(firsttimesort[[H]][,21])
# 
#     #reset k index to 1 for new serial number/lot - should do this automatically with every change in level
# 
#     for (k in 1:max(c(timecountL*2, timecountN*2, timecountH*2))) {
#       #reset k if < 1?
#       #if(k<1){k=1}
#       
#       #Low Date check for time skip
#       #Check if same row date is =, skip if > < or NA
#       
#       #Low to Normal check
#       if(if_else(firsttimesort[[L]][k,"Date"] == firsttimesort[[N]][k,"Date"], TRUE, FALSE, missing = FALSE)){
#         #Normal time skip
#         if (if_else(firsttimesort[[L]][k,21] - firsttimesort[[N]][k,21] > timediff, TRUE, FALSE, missing = FALSE)){
#           firsttimesort[[L]] <- InsertRow(firsttimesort[[L]], NArow, RowNum = k)}}
#       
#       #Low to High check
#       if(if_else(firsttimesort[[L]][k,"Date"] == firsttimesort[[H]][k,"Date"], TRUE, FALSE, missing = FALSE)){
#         #High time skip
#         if (if_else(firsttimesort[[L]][k,21] - firsttimesort[[H]][k,21] > timediff, TRUE, FALSE, missing = FALSE)){
#           firsttimesort[[L]] <- InsertRow(firsttimesort[[L]], NArow, RowNum = k)}}
#       
#      
#       #Normal Date check for time skip
#       #Check if same row date is =, skip if > < or NA
#       
#       #Normal to Low check
#       if(if_else(firsttimesort[[N]][k,"Date"] == firsttimesort[[L]][k,"Date"], TRUE, FALSE, missing = FALSE)){
#         #Low time skip
#         if (if_else(firsttimesort[[N]][k,21] - firsttimesort[[L]][k,21] > timediff, TRUE, FALSE, missing = FALSE)){
#           firsttimesort[[N]] <- InsertRow(firsttimesort[[N]], NArow, RowNum = k)}}
#       
#         #Normal to High check
#         if(if_else(firsttimesort[[N]][k,"Date"] == firsttimesort[[H]][k,"Date"], TRUE, FALSE, missing = FALSE)){
#             #High time skip
#           if (if_else(firsttimesort[[N]][k,21] - firsttimesort[[H]][k,21] > timediff, TRUE, FALSE, missing = FALSE)){
#             firsttimesort[[N]] <- InsertRow(firsttimesort[[N]], NArow, RowNum = k)}}
#           
#         
#       #High Date check for time skip
#       #Check if same row date is =, skip if > < or NA
#       
#       #High to Low check
#       if(if_else(firsttimesort[[H]][k,"Date"] == firsttimesort[[L]][k,"Date"], TRUE, FALSE, missing = FALSE)){
#         #Low time skip
#         if (if_else(firsttimesort[[H]][k,21] - firsttimesort[[L]][k,21] > timediff, TRUE, FALSE, missing = FALSE)){
#           firsttimesort[[H]] <- InsertRow(firsttimesort[[H]], NArow, RowNum = k)}}
#         
#         #High to Normal check
#         if(if_else(firsttimesort[[H]][k,"Date"] == firsttimesort[[N]][k,"Date"], TRUE, FALSE, missing = FALSE)){
#           #Normal time skip
#           if (if_else(firsttimesort[[H]][k,21] - firsttimesort[[N]][k,21] > timediff, TRUE, FALSE, missing = FALSE)){
#             firsttimesort[[H]] <- InsertRow(firsttimesort[[H]], NArow, RowNum = k)}}
#           
#                   
#       ####Check if row date directly above is equal (or why bother skipping a line)
#       #only run if is k >1
#       # if(k > 1){
#       #   
#       #   #Low self date check, is date directly above same?
#       #   if (if_else(firsttimesort[[L]][k,"Date"] == firsttimesort[[L]][k-1,"Date"], TRUE, FALSE, missing = FALSE)){
#       #     #Low self time skip
#       #     if (if_else(firsttimesort[[L]][k,21] - firsttimesort[[L]][k-1,21] > timediff, TRUE, FALSE, missing = FALSE)){
#       #       firsttimesort[[L]] <- InsertRow(firsttimesort[[L]], NArow, RowNum = k) k=k-1 & next}}
#       #     
#       #     #Normal self date check, is date directly above same?
#       #     if(if_else(firsttimesort[[N]][k,"Date"] == firsttimesort[[N]][k-1,"Date"], TRUE, FALSE, missing = FALSE)){
#       #       #Normal self time skip
#       #       if (if_else(firsttimesort[[N]][k,21] - firsttimesort[[N]][k-1,21] > timediff, TRUE, FALSE, missing = FALSE)){
#       #         firsttimesort[[N]] <- InsertRow(firsttimesort[[N]], NArow, RowNum = k) k=k-1 & next}}
#       #     
#       #     #High self date check, is date directly above same?
#       #     if(if_else(firsttimesort[[H]][k,"Date"] == firsttimesort[[H]][k-1,"Date"], TRUE, FALSE, missing = FALSE)){
#       #       firsttimesort[[H]] <- InsertRow(firsttimesort[[H]], NArow, RowNum = k)
#       #     #High self time skip
#       #      if (if_else(firsttimesort[[H]][k,21] - firsttimesort[[H]][k-1,21] > timediff, TRUE, FALSE, missing = FALSE)){
#       #     firsttimesort[[N]] <- InsertRow(firsttimesort[[N]], NArow, RowNum = k) k=k-1 & next}}
#       #     
#       #   }
#       } #k bracket close
#   
#    #i usn/level of loop
#    
#    #recount table length
#    repeat{
#      timecountL <- length(firsttimesort[[L]][,21])
#      timecountN <- length(firsttimesort[[N]][,21])
#      timecountH <- length(firsttimesort[[H]][,21])
#    
#      #Add rows to bottom of table until table length is equal
#      #low check
#      if (timecountL < max(timecountL, timecountN, timecountH)){
#    
#        firsttimesort[[L]] <- rbind(firsttimesort[[L]], NArow)
#      }
#    
#      #normal check
#      if (timecountN < max(timecountL, timecountN, timecountH)){
#    
#        firsttimesort[[N]] <- rbind(firsttimesort[[N]], NArow)
#      }
#    
#      #high check
#      if (timecountH < max(timecountL, timecountN, timecountH)){
#    
#        firsttimesort[[H]] <- rbind(firsttimesort[[H]], NArow)
#      }
#    
#      #break when all column lengths are equal
#      if(timecountL == timecountN & timecountN == timecountH){break}
# 
#   } #repeat loop bracket
#     
# }  #i bracket close
#   
#  } #j bracket close
# 
# 
# timeselfcheck <- firsttimesort
# 
# 
# #SEPARATE SELF TIME CHECK LOOP
# 
# for (j in 1:length(usn)){
#   
#   for (i in 1:length(ulevel)){
#     
#     #assign search "level sn" change lot to L N H for each element
#     L <- paste(usn[j], "L", ulevel[i], sep = "_")
#     N <- paste(usn[j], "N", ulevel[i], sep = "_")
#     H <- paste(usn[j], "H", ulevel[i], sep = "_")
#     
#     ##check if sn/qc combo exists in list.qc
#     if (!exists(L, where = timeselfcheck)) {next}
#     if (!exists(N, where = timeselfcheck)) {next}
#     if (!exists(H, where = timeselfcheck)) {next}
#     
#     #compare time rows to set index to longest row
#     #Time = col 21
#     timecountfinalL <- length(timeselfcheck[[L]][,21])
#     timecountfinalN <- length(timeselfcheck[[N]][,21])
#     timecountfinalH <- length(timeselfcheck[[H]][,21])
#     
#     #reset k index to 1 for new serial number/lot - should do this automatically with every change in level
#     
#     for (k in 2:max(c(timecountL*2, timecountN*2, timecountH*2))) {
#       #reset k if < 1?
#       #if(k<1){k=1}
#       
#       #Check if row date directly above is equal
#       #only run if is k >1
#       # if(k > 1){
#          
#          #Low self date check, is date directly above same?
#          if (if_else(timeselfcheck[[L]][k,"Date"] == timeselfcheck[[L]][k-1,"Date"], TRUE, FALSE, missing = FALSE)){
#            #Low self time skip
#            if (if_else(timeselfcheck[[L]][k,21] - timeselfcheck[[L]][k-1,21] > timediff, TRUE, FALSE, missing = FALSE)){
#              timeselfcheck[[L]] <- InsertRow(timeselfcheck[[L]], NArow, RowNum = k)}}
#            
#            #Normal self date check, is date directly above same?
#            if(if_else(timeselfcheck[[N]][k,"Date"] == timeselfcheck[[N]][k-1,"Date"], TRUE, FALSE, missing = FALSE)){
#              #Normal self time skip
#              if (if_else(timeselfcheck[[N]][k,21] - timeselfcheck[[N]][k-1,21] > timediff, TRUE, FALSE, missing = FALSE)){
#                timeselfcheck[[N]] <- InsertRow(timeselfcheck[[N]], NArow, RowNum = k)}}
#            
#            #High self date check, is date directly above same?
#            if(if_else(timeselfcheck[[H]][k,"Date"] == timeselfcheck[[H]][k-1,"Date"], TRUE, FALSE, missing = FALSE)){
#              timeselfcheck[[H]] <- InsertRow(timeselfcheck[[H]], NArow, RowNum = k)
#            #High self time skip
#             if (if_else(timeselfcheck[[H]][k,21] - timeselfcheck[[H]][k-1,21] > timediff, TRUE, FALSE, missing = FALSE)){
#            timeselfcheck[[N]] <- InsertRow(timeselfcheck[[N]], NArow, RowNum = k)}}
#            
#     # } if k >1 bracket
#     } #k bracket close
#     
#     #i usn/level of loop
#     
#     #recount table length
#     repeat{
#       timecountfinalL <- length(timeselfcheck[[L]][,21])
#       timecountfinalN <- length(timeselfcheck[[N]][,21])
#       timecountfinalH <- length(timeselfcheck[[H]][,21])
#       
#       #Add rows to bottom of table until table length is equal
#       #low check
#       if (timecountfinalL < max(timecountfinalL, timecountfinalN, timecountfinalH)){
#         
#         timeselfcheck[[L]] <- rbind(timeselfcheck[[L]], NArow)
#       }
#       
#       #normal check
#       if (timecountfinalN < max(timecountfinalL, timecountfinalN, timecountfinalH)){
#         
#         timeselfcheck[[N]] <- rbind(timeselfcheck[[N]], NArow)
#       }
#       
#       #high check
#       if (timecountfinalH < max(timecountfinalL, timecountfinalN, timecountfinalH)){
#         
#         timeselfcheck[[H]] <- rbind(timeselfcheck[[H]], NArow)
#       }
#       
#       #break when all column lengths are equal
#       if(timecountfinalL == timecountfinalN & timecountfinalN == timecountfinalH){break}
#       
#     } #repeat loop bracket
#     
#   }  #i bracket close
#   
# } #j bracket close
#       
# 
# 
# #Change time back to chacter and remove sysdate that was added.
# sysdate <- as.character(Sys.Date())
# 
# for (i in 1:length(timeselfcheck)) {
#   
# #date class change to character
# timeselfcheck[[i]][,21] <- as.character(timeselfcheck[[i]][,21])
# 
# timeselfcheck[[i]][,21] <- str_remove(timeselfcheck[[i]][,21], sysdate)
# #remove sysdate that was added
# 
# }
# 
# #Second Date sort after time sort:
# finaldatesort <- timeselfcheck
# 
# 
# for (j in 1:length(usn)){
#   
#   #QClevel counter
#   for (i in 1:length(ulevel)){
#     
#     #assign search "level sn" change lot to L N H for each element
#     L <- paste(usn[j], "L", ulevel[i], sep = "_")
#     N <- paste(usn[j], "N", ulevel[i], sep = "_")
#     H <- paste(usn[j], "H", ulevel[i], sep = "_")
#     
#     #check if sn/qc combo exists in list.qc
#     if (!exists(L, where = finaldatesort)) {next}
#     if (!exists(N, where = finaldatesort)) {next}
#     if (!exists(H, where = finaldatesort)) {next}
#     
#     #compare date rows to set index to longest row
#     #unable to change iteration number inside loop, set it to arbitrarily high number max datecount *3
#     finaldatecountL <- length(finaldatesort[[L]][["Date"]])
#     finaldatecountN <- length(finaldatesort[[N]][["Date"]])
#     finaldatecountH <- length(finaldatesort[[H]][["Date"]])
#     
#     
#     #reset k index to 1 for new serial number/lot - should do this automatically with every change in level
#     #k = 1
#     for (k in 1:max(finaldatecountL*2, finaldatecountN*2, finaldatecountH*2)) {
#       
#       #Low to Normal check
#       #new table edited
#       if (if_else(finaldatesort[[L]][k,"Date"] > finaldatesort[[N]][k,"Date"], TRUE, FALSE, missing = FALSE)){finaldatesort[[L]] <- InsertRow(finaldatesort[[L]], NArow, RowNum = k)
#       }
#       
#       #Low to High check                                                   
#       #new table edited
#       if (if_else(finaldatesort[[L]][k,"Date"] > finaldatesort[[H]][k,"Date"], TRUE, FALSE, missing = FALSE)){finaldatesort[[L]] <- InsertRow(finaldatesort[[L]], NArow, RowNum = k)
#       }
#       
#       #Normal to Low check                                                  
#       #new table edited
#       if (if_else(finaldatesort[[N]][k,"Date"] > finaldatesort[[L]][k,"Date"], TRUE, FALSE, missing = FALSE)){finaldatesort[[N]] <- InsertRow(finaldatesort[[N]], NArow, RowNum = k)
#       }
#       
#       #Normal to High check  
#       #new table edited
#       if (if_else(finaldatesort[[N]][k,"Date"] > finaldatesort[[H]][k,"Date"], TRUE, FALSE, missing = FALSE)){finaldatesort[[N]] <- InsertRow(finaldatesort[[N]], NArow, RowNum = k)
#       }
#       
#       #High to Low Check
#       #new table edited
#       if (if_else(finaldatesort[[H]][k,"Date"] > finaldatesort[[L]][k,"Date"], TRUE, FALSE, missing = FALSE)){finaldatesort[[H]] <- InsertRow(finaldatesort[[H]], NArow, RowNum = k)
#       }
#       
#       #High to Normal Check
#       #new table edited
#       if (if_else(finaldatesort[[H]][k,"Date"] > finaldatesort[[N]][k,"Date"], TRUE, FALSE, missing = FALSE)){finaldatesort[[H]] <- InsertRow(finaldatesort[[H]], NArow, RowNum = k)
#       }
#       
#       ###Less than check ### Moves the other file down that is not equal but not less than (so must be greater than)
#       ##This code not needed after final datecount length check and coalesce
# # 
# #         #Normal to Low check
# #         #new table edited
# #         if (if_else(finaldatesort[[N]][k,"Date"] < finaldatesort[[L]][k,"Date"], TRUE, FALSE, missing = FALSE)){finaldatesort[[L]] <- InsertRow(finaldatesort[[L]], NArow, RowNum = k)
# #          k=k-1}
# # 
# #         #High to Low check
# #         #new table edited
# #         if (if_else(finaldatesort[[H]][k,"Date"] < finaldatesort[[L]][k,"Date"], TRUE, FALSE, missing = FALSE)){finaldatesort[[L]] <- InsertRow(finaldatesort[[N]], NArow, RowNum = k)
# #          k=k-1}
# # 
# #         #Low to Normal check
# #         #new table edited
# #         if (if_else(finaldatesort[[L]][k,"Date"] < finaldatesort[[N]][k,"Date"], TRUE, FALSE, missing = FALSE)){finaldatesort[[N]] <- InsertRow(finaldatesort[[L]], NArow, RowNum = k)
# #          k=k-1}
# # 
# #         #High to Normal check
# #         #new table edited
# #         if (if_else(finaldatesort[[H]][k,"Date"] < finaldatesort[[N]][k,"Date"], TRUE, FALSE, missing = FALSE)){finaldatesort[[N]] <- InsertRow(finaldatesort[[N]], NArow, RowNum = k)
# #          k=k-1}
# # 
# #         #High to Low Check
# #         #new table edited
# #         if (if_else(finaldatesort[[L]][k,"Date"] < finaldatesort[[H]][k,"Date"], TRUE, FALSE, missing = FALSE)){finaldatesort[[H]] <- InsertRow(finaldatesort[[H]], NArow, RowNum = k)
# #          k=k-1}
# # 
# #         #High to Normal Check
# #         #new table edited
# #         if (if_else(finaldatesort[[N]][k,"Date"] < finaldatesort[[H]][k,"Date"], TRUE, FALSE, missing = FALSE)){finaldatesort[[H]] <- InsertRow(finaldatesort[[H]], NArow, RowNum = k)
# #          k=k-1}
# 
# 
#     }
#     
#     #j sn/level loop after sorting - Make tables same length (add NA if necessary), then merge dates for L N H here  each lot/sn after date sort
#     
#     #recount table length
#     repeat{
#       finaldatecountL <- length(finaldatesort[[L]][["Date"]])
#       finaldatecountN <- length(finaldatesort[[N]][["Date]])
#       finaldatecountH <- length(finaldatesort[[H]][["Date]])
#       
#       #Add rows to bottom of table until table length is equal
#       #low check
#       if (finaldatecountL < max(finaldatecountL, finaldatecountN, finaldatecountH)){
#         
#         finaldatesort[[L]] <- rbind(finaldatesort[[L]], NArow)
#       }
#       
#       #normal check
#       if (finaldatecountN < max(finaldatecountL, finaldatecountN, finaldatecountH)){
#         
#         finaldatesort[[N]] <- rbind(finaldatesort[[N]], NArow)
#       }
#       
#       #high check
#       if (finaldatecountH < max(finaldatecountL, finaldatecountN, finaldatecountH)){
#         
#         finaldatesort[[H]] <- rbind(finaldatesort[[H]], NArow)
#       }
#       
#       #break when all column lengths are equal
#       if(finaldatecountL == finaldatecountN & finaldatecountN == finaldatecountH){break}
#     }
# 
#     #if (if_else(finaldatesort[[L]][["Date]] == finaldatesort[[N]][["Date"]] & finaldatesort[[H]][["Date"]] == finaldatesort[[L]][["Date]], TRUE, FALSE, missing = TRUE)){
# 
#     #unlist to vectorize, retains Date type but no longer a dataframe
#     Lfinalunlist <- unlist(finaldatesort[[L]][["Date"]])
#     Nfinalunlist <- unlist(finaldatesort[[N]][["Date"]])
#     Hfinalunlist <- unlist(finaldatesort[[H]][["Date"]])
# 
#     #coalesce to merge dates, only if vectors lengths are equal
#     
#     #replace old dates or set equal to new
#     finaldatesort[[L]][["Date"]] <- as.data.frame(coalesce(Lfinalunlist, Nfinalunlist, Hfinalunlist))
#     finaldatesort[[N]][["Date"]] <- as.data.frame(coalesce(Lfinalunlist, Nfinalunlist, Hfinalunlist))
#     finaldatesort[[H]][["Date"]] <- as.data.frame(coalesce(Lfinalunlist, Nfinalunlist, Hfinalunlist))
# 
#     #} for if statement if used, tables are already checked for length
#   }
# } 
# 
# 
# # #Remove date from rows that are all NA except date then coalesce
# # for (j in 1:length(usn)){
# #   
# #   #QClevel counter
# #   for (i in 1:length(ulevel)){
# #     
# #     #assign search "level sn" change lot to L N H for each element
# #     L <- paste(usn[j], "L", ulevel[i], sep = "_")
# #     N <- paste(usn[j], "N", ulevel[i], sep = "_")
# #     H <- paste(usn[j], "H", ulevel[i], sep = "_")
# #     
# #     #check if sn/qc combo exists in list.qc
# #     if (is.null(finaldatesort[[L]])) {next}
# #     if (is.null(finaldatesort[[N]])) {next}
# #     if (is.null(finaldatesort[[H]])) {next}
# #     
# #     for (k in 1:length(max(finaldatesort[[L]][["Date"]],finaldatesort[[N]][["Date"]],finaldatesort[[L]][["Date"]]))) {
# #       
# #     if(sum(is.na(finaldatesort[[L]][k,])) == 20){ finaldatesort[[L]][k,] <- NArow}
# #     
# #       if(sum(is.na(finaldatesort[[N]][k,])) == 20){ finaldatesort[[N]][k,] <- NArow}
# #       
# #       if(sum(is.na(finaldatesort[[H]][k,])) == 20){ finaldatesort[[H]][k,] <- NArow}
# #     
# #   }}}



#Add sn, lot, level, and rename date columns to table
for (i in 1:length(list.qc)) {
  list.qc[[i]] <- cbind(sn[i], level[i], lot[i], list.qc[[i]][2:length(list.qc[[i]])])
  
  namefix  <- c("SN", "Level", "Lot", names(list.qc[[i]][4:length(list.qc[[i]])]))
  
  colnames(list.qc[[i]]) <- namefix
  
  #fixed in earlier loop
  #date name fix
  #colnames(list.qc[[i]])[grep(x = colnames(list.qc[[i]]), pattern = "Date")] <- "Date"
  
}



#wrap all this into 1 for loop or just use purr:map on each
#Add sn, lot, level, and comment columns to table
for (i in 1:length(firstdatesort)) {
  firstdatesort[[i]] <- cbind(sn[i], level[i], lot[i], firstdatesort[[i]][2:length(firstdatesort[[i]])])
  
  namefix  <- c("SN", "Level", "Lot", names(firstdatesort[[i]][4:length(firstdatesort[[i]])]))
  
  colnames(firstdatesort[[i]]) <- namefix
  
  #date name fix
  #colnames(firstdatesort[[i]])[grep(x = colnames(firstdatesort[[i]]), pattern = "Date")] <- "Date"
  
}




# #Add sn, lot, level, and comment columns to table
# for (i in 1:length(firsttimesort)) {
#   firsttimesort[[i]] <- cbind(sn[i], level[i], lot[i], firsttimesort[[i]][2:length(firsttimesort[[i]])])
#   
#   namefix  <- c("SN", "Level", "Lot", names(firsttimesort[[i]][4:length(firsttimesort[[i]])]))
#   
#   colnames(firsttimesort[[i]]) <- namefix
#
#   #date name fix
#   colnames(firsttimesort[[i]])[grep(x = colnames(firsttimesort[[i]]), pattern = "Date")] <- "Date"

#}

# #Add sn, lot, level, and comment columns to table
# for (i in 1:length(timeselfcheck)) {
#   timeselfcheck[[i]] <- cbind(sn[i], level[i], lot[i], timeselfcheck[[i]][2:length(timeselfcheck[[i]])])
#   
#   namefix  <- c("SN", "Level", "Lot", names(timeselfcheck[[i]][4:length(timeselfcheck[[i]])]))
#   
#   colnames(timeselfcheck[[i]]) <- namefix
#
#   #date name fix
#   colnames(timeselfcheck[[i]])[grep(x = colnames(timeselfcheck[[i]]), pattern = "Date")] <- "Date"

# }



#Create tables ready for publishing/writing.

#in the future change so LNH are just put in another list?

#Add "Lot #" to first row of Comment column
for (i in 1:length(firstdatesort)) {
  
  firstdatesort[[i]][["Comments"]][[1]] <- paste( "Lot", firstdatesort[[i]][["Lot"]][[1]])
  
}


#Add number on to replicate dates as a new column "ShortDate"
for (i in 1:length(firstdatesort)) {
   
  firstdatesort[[i]][["ShortDate"]]  <- getanID(firstdatesort[[i]][["Date"]])[[".id"]]
  
  firstdatesort[[i]][["ShortDate"]] <- paste0(firstdatesort[[i]][["Date"]], " #", firstdatesort[[i]][["ShortDate"]])

  #overwrite any values that are not #1 with just the number
  firstdatesort[[i]][["ShortDate"]][grep("#1$", firstdatesort[[i]][["ShortDate"]], invert = TRUE)] <- 
    str_sub(firstdatesort[[i]][["ShortDate"]][grep("#1$", firstdatesort[[i]][["ShortDate"]], invert = TRUE)], start = -3, end = -1)
   
 }


#use purr map here to set date format faster
#set date class
#set time/period class
for (i in 1:length(firstdatesort)){
  
  #set date column to Date class for comparison later.
  #firstdatesort[[i]][["Date"]] <- mdy(firstdatesort[[i]][["Date"]])
  
  #set time column class
  firstdatesort[[i]][[grep("Time", colnames(firstdatesort[[i]]))]] <- hms(firstdatesort[[i]][[grep("Time", colnames(firstdatesort[[i]]))]], quiet = TRUE)
  
  #create new datetime column for graphing later
  firstdatesort[[i]][grep("datetime", colnames(firstdatesort[[i]]))] <- ymd_hms(paste(mdy(firstdatesort[[i]][["Date"]], quiet = TRUE),
                                                    firstdatesort[[i]][[grep("Time", colnames(firstdatesort[[i]]))]], sep = " "), quiet = TRUE)
  
}


#Create LNH final tables.

for (j in 1:length(usn))
  for (i in 1:length(ulevel)) {{
    
    #bind search "level sn" change lot to L N H for each element
    L <- paste(usn[j],"L", ulevel[i], sep = "_")
    N <- paste(usn[j],"N", ulevel[i], sep = "_")
    H <- paste(usn[j],"H", ulevel[i], sep = "_")
    
    #check if object exists in firstdatesort
    if (!exists(L, where = firstdatesort)) {next}
    if (!exists(N, where = firstdatesort)) {next} 
    if (!exists(H, where = firstdatesort)) {next} 
    
    assign(paste("LNH", usn[j],"_", ulevel[i], sep=""), cbindX(
      firstdatesort[[L]], 
      firstdatesort[[N]],
      firstdatesort[[H]])[finalcols])
    
    
  }}

#AVQC units, disabled at Lihjen's request
#avunits <- c("10^9/L", "10^9/L", "%", "10^9/L", "%", "10^9/L", "%", "10^9/L", "g/dL", "%", "fL", "pg", "g/dL", "fL","10^9/L", "fL", "%", "%")


### MoVED AVFILES OBJECT UP FOR FASTER STOP IF NO QC FILES ###

#read tables
avalues <- lapply(avfiles, read.table, header=FALSE, fill = TRUE)

#transpose tables
avalues <- lapply(avalues, function(x) t(x))

#fix after transpose
avalues <- lapply(avalues, as.data.frame)

#name tables
avnames <- paste(str_sub(avfiles, -3, -1), str_extract(avfiles, pattern = "[:digit:][:digit:][:digit:][:digit:][:digit:]?[:digit:]?"), sep = "")
names(avalues) <- avnames

#fix again, convert from matrix to data.frame
avalues <- lapply(avalues, as.data.frame)
  
#name first rows Mean and SD
av <- c("Target", "Limit")

  #Empty date and Time just for naming
  #DandT <- list("", "")
  
  #QC header columns 1    2     3       4     5     6     7     8       9     10      11      12    13      14    15      16      17    18
  avheaders <- c("WBC", "LYM","%L", "MID", "%M", "GRAN", "%G", "RBC", "HGB", "HCT", "MCV", "MCH", "MCHC", "RDW", "PLT", "MPV", "PCT", "PDW")

  #assign/promote column names and add "Low", "Normal", "High"

  for (i in 1:length(avalues)) {
    
    #Add extra columns Date and Time so they get LL NL HL
    #avalues[[i]] <- cbind(avalues[[i]], DandT)
    
    
    #Low
    if (grepl("*L", names(avalues[i])))
    {names(avalues[[i]]) <- c("Date", "Lot", NA, paste(avheaders, "LL", sep = "_"))}
    
    #Normal
    if (grepl("*N", names(avalues[i])))
    {names(avalues[[i]]) <- c("Date", "Lot", NA, paste(avheaders, "NL", sep = "_"))}
    
    #High
    if (grepl("*H", names(avalues[i])))
    {names(avalues[[i]]) <- c("Date", "Lot", NA, paste(avheaders, "HL", sep = "_"))}
    
  
  }
  
#assign names and make LNH object
#empty date = 1, empty sn =2, RBC = 8, MCV = 11, PLT = 15, 
  
#sort columns to match QC columns, remove extra row and add Target and Mean to Column 1

  for (i in 1:length(avalues)) {
    
    #Put lot in second row with expiration
    #Put expiration in third row as date and format to %m/%d/%y
    avalues[[i]][,2] <- as.character(avalues[[i]][,2])
    avalues[[i]][2,2] <- sub("[[:alpha:]]", "",avalues[[i]][1,2])
    avalues[[i]][3,2] <- format(as.Date(avalues[[i]][1,3], format = "%d/%m/%y"), format = "%m/%d/%y")
    
    
    #sort columns to match QC columns
    avalues[[i]] <- avalues[[i]][,c(1, 4, 9, 10, 5, 6, 7, 8, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21,2)]
    
  }  

  #remove extra row

  for (i in 1:length(avalues)) {
    
          avalues[[i]] <- avalues[[i]][2:3,]
  }

  #add av (Target and Mean) to 

  #use lapply to make each only include row 2 and another lapply for making numeric
  #make numeric instead of factors and remove first row.
  for (i in 1:length(avalues)) {
  
    avalues[[i]][,1] <- av
    
    avalues[[i]][,2:19] <- as.numeric(as.character(unlist(avalues[[i]][,2:19])))
   
      #with avunits, disabled at Lihjen's request   
  #avalues[[i]] <- rbind(as.numeric(as.character(unlist(avalues[[i]][2,]))),
      #as.numeric(as.character(unlist(avalues[[i]][3,]))), avunits)
  
    #Add NA bottom row to space out in spreadsheet
    avalues[[i]][c(nrow(avalues[[i]])+1),] <- NA
  }


#KEY NOW SAME AS QC NOW
#  1   2       3     4       5     6     7        8     9    10     11      12     13    14      15     16     17      18    19      20       21
# lot "WBC", "GRAN", "%G", "LYM", "%L", "MID", "%M", "RBC", "HGB", "HCT", "MCV", "MCH", "MCHC", "RDW", "PLT", "MPV", "PCT", "PDW", "Date", "Time"



#For writing QC values to lot A, B, etc. sheets, level without number at end
qclevel <- sub("[[:alpha:]]$", "",ulevel)


for (i in 1:length(qclevel))
{
  #check if that qc lot exists
  if (!exists(paste("QCL", qclevel[i], sep = ""), where = avalues)){next}
  if (!exists(paste("QCN", qclevel[i], sep = ""), where = avalues)){next}
  if (!exists(paste("QCH", qclevel[i], sep = ""), where = avalues)){next}
  
  
  
  assign(paste("QCLNH", qclevel[i], sep=""), cbind(
    avalues[[paste("QCL", qclevel[i], sep = "")]], 
    avalues[[paste("QCN", qclevel[i], sep = "")]],
    avalues[[paste("QCH", qclevel[i], sep = "")]])[avcols]
  )
}


#Create centered text style
centered <- createStyle(halign = "center")
  
#Create red text style
red <- createStyle(fontColour = "red")

#Create bold text style
bold <- createStyle(textDecoration = "bold")

#Create  style for 1 decimal place values WBC LYM %L MID %M GRAN %G HGB HCT MCV MCH MCHC
onedigit <- createStyle(numFmt = "0.0")

##Create  style for 2 decimal place values RBC
twodigit <- createStyle(numFmt = "0.00")

#need to create blank empty workbook to change name
wb <- createWorkbook()



#Check for or create subfolder for spreadsheet files
folder <- paste(getwd(), "/R Generated Spreadsheets", sep = "")

if(!dir.exists(folder)){
  
  #create directory to save spreadsheets in
  dir.create(folder)
}


#qc level counter, including letter QC
for (i in 1:length(ulevel)){
  
  #Check if QC lot exists
  finalqc <- paste("QCLNH", qclevel[i], sep = "")

  # ##check if qc lot exists or skip
   if (!exists(finalqc)) {next}
  
  #Skip if file exists, go to next level of loop so time isn't wasted making graphs and wb_QCLNH object
  if (file.exists(paste("E18 All QC for lot ", ulevel[i], ".xlsx", sep = ""))) {next}
    
  #Create blank workbook for each valid QC level
  
  assign(paste("wb_QCLNH", ulevel[i], sep = ""), copyWorkbook(wb)) 

  
  #sn counter
  for (j in 1:length(usn)){

    #assign search "level sn" change lot to L N H for each element
    finaldata <- paste("LNH", usn[j],"_", ulevel[i], sep="")
    finalqc <- paste("QCLNH", qclevel[i], sep = "")
    finalwb <- paste("wb_QCLNH", ulevel[i], sep = "")
    
    ##check if LNH combo exists or skip
    if (!exists(finaldata)) {next}
    
    #Create a sheet for every sn or do through writing qc assay values if possible
    addWorksheet(get(finalwb), usn[j])
    
  #Write QC assay values to workbook object and create worksheets for every instrument?
  writeData(get(finalwb), usn[j], get(finalqc), colNames = TRUE, startRow = assaystart) 
  
  
  #Use WriteData to write QC data to already created sheet
  writeData(get(finalwb), usn[j], get(finaldata),
            startRow = c(datastart), colNames = TRUE)
    
  #Add 1 extra digits for  WBC LYM %L MID %M GRAN %G HGB HCT MCV MCH MCHC
  addStyle(get(finalwb), usn[j], onedigit, 
           rows = assaystart:c(nrow(get(finaldata))+datastart),
           cols = grep("WBC|LYM|%L|MID|%M|GRAN|%G|HGB|HCT|MCV|MCH|MCHC", finalcols), gridExpand = TRUE, stack = TRUE)
           
  #Add 2 extra digits for RBC
  addStyle(get(finalwb), usn[j], twodigit, 
           rows = assaystart:c(nrow(get(finaldata))+datastart),
           cols = grep("RBC", finalcols), gridExpand = TRUE, stack = TRUE)
  
  #Add centered formatting for finalcols, disable if it works through formattable
  addStyle(get(finalwb), usn[j], centered, 
           rows = assaystart:c(nrow(get(finaldata))+datastart),
           cols = 1:length(finalcols), gridExpand = TRUE, stack = TRUE)
  
  #Add bold assay value colnames
  addStyle(get(finalwb), usn[j], bold, 
           rows = c(assaystart, datastart), cols = 1:length(finalcols), gridExpand = TRUE, stack = TRUE)
  
  #Add permanent red formatting in the future withto cells that are OOR starting in row 7
  #addStyle(get(finalwb), usn[j], red, rows = x, cols = y, gridExpand = TRUE, stack = TRUE)
  
  #Use conditional formatting for now, could test "red" addstyle but probably too slow
  #for loop to add to each column
  for (k in 1:length(valuecols)) {

    #less than rule
    conditionalFormatting(get(finalwb), usn[j], cols = k+1, 
                          rows = datastart+1:nrow(get(finaldata)),
                          style = red, type = "expression",                    #c(lower limit, upper limit)
                          rule = paste("<", c(get(paste("QCLNH", qclevel[i], sep = ""))[1,valuecols[k]] - get(paste("QCLNH", qclevel[i], sep = ""))[2,valuecols[k]])))
    #greater than rule
    conditionalFormatting(get(finalwb), usn[j], cols = k+1, 
                          rows = datastart+1:nrow(get(finaldata)),
                          style = red, type = "expression",                    #c(lower limit, upper limit)
                          rule = paste(">", c(get(paste("QCLNH", qclevel[i], sep = ""))[1,valuecols[k]] + get(paste("QCLNH", qclevel[i], sep = ""))[2,valuecols[k]])))
    
  } #end k loop
  
  #set or reset r row counter for new workbook
  r = 2
  
  #switch to turn graphs on or off:
  if (graphs) {
    
  
  
  #Add timeseries ggplots of qc data one at a time for every variable, make plot then insertplot into wb object.
  #use r for row # to space out graphs
  for (m in 1:length(valuecols)) {
   
                                          # get datetime x-axis, choose LL, NL, or HL
     plot <- ggplot(data = get(finaldata), aes_string(x= paste("datetime", str_sub(valuecols[m], start = -2, end = -1), sep = "_"), y = valuecols[m])) + geom_line(color = "blue", na.rm = TRUE)
     
     #Add title to graph
     plot <- plot + ggtitle(paste("Emerald 18 SN", usn[j], "Lot", ulevel[i], valuecols[m])) + theme_bw()
     
     #Add Target
     plot <- plot + geom_hline(yintercept =  get(paste("QCLNH", qclevel[i], sep = ""))[1,valuecols[m]],
                               "line", color = "red")
     
     #Add upper limit
     plot <- plot + geom_hline(yintercept =  get(paste("QCLNH", qclevel[i], sep = ""))[1,valuecols[m]] + get(paste("QCLNH", qclevel[i], sep = ""))[2,valuecols[m]],
                               "line", color = "orange")
    
     #Add lower limit
     plot <- plot + geom_hline(yintercept =  get(paste("QCLNH", qclevel[i], sep = ""))[1,valuecols[m]] - get(paste("QCLNH", qclevel[i], sep = ""))[2,valuecols[m]],
                               "line", color = "orange")
     
     #to show plot in viewer for insertplot function
     print(plot)
     
    insertPlot(get(finalwb), usn[j], startCol = length(finalcols)+3, startRow =  r)
    
    r= r + 20
  }                                                                                                                             #this code written by JIT
  } #end if statement on/off switch
  
  } #j usn bracket close
  
  
  #Write wb object to file and name after QC level
  #Use tryCatch to go to next if file exists
  #might be able to just use try() here instead? unsure of how try functions
  tryCatch(expr =  saveWorkbook(get(finalwb), 
                        paste( folder, "/E18 All QC for lot ", ulevel[i], ".xlsx", sep = ""), overwrite = FALSE),
           finally = next)
            #go to next # in loop if file exists
  
  } #i ulevel bracket close




##########################                 Continuous workbooks for each SN            ########################################################


#Find all QC by SN in order and bind it to one huge table
#Also bind QC based on lots in the list
for (i in 1:length(usn)) {
  
  datalist <- apropos(paste("^LNH", usn[i],"_", sep=""))
  
  #combine all tables of that SN  
  alldata <- lapply(datalist, get)

  #Assign new name to combine data table
  assign(paste("alldata", usn[i], sep = ""), bind_rows(alldata))
    
  
  #Remove LNH[SN] and keep lot #, assign name for use in conditional format loop
  allqc <- sub(str_flatten(paste( "LNH", usn[i], "_", sep = ""),"|"), "", datalist)
  
  #Remove letter from QC lot number, same as qclevel object for each sn
  allqc <- sub("[[:alpha:]]$", "", allqc)
  
  
  #Get QC table names and get objects for binding
  allqc <- lapply(paste("QCLNH", allqc, sep = ""), get)
  
  #Check if QC lot exists
  #finalqc <- paste("QCLNH", qclevel[i], sep = "")
  
  #Get all QC files and bind
  assign(paste("allqc", usn[i], sep = ""), bind_rows(allqc))
  
}


#Create continuous workbook for each SN


#usn counter only, no u/qclevel counter needed
for (i in 1:length(usn)){

  #assign search "level sn" change lot to L N H for each element
  allfinaldata <- paste("alldata", usn[i], sep="")
  allfinalqc <- paste("allqc", usn[i], sep = "")
  allfinalwb <- paste("wb_alldata", usn[i], sep = "")
  
  #datastart changes based on length of qc assayvalues for continuous workbook
  alldatastart <- nrow(get(allfinalqc))+assaystart+datastart/2
  
  #Skip if file exists, go to next level of loop so time isn't wasted making graphs and wb_QCLNH object
  if (file.exists(paste("E18 All QC for SN ", usn[i], ".xlsx", sep = ""))) {next}

  #Create blank workbook for each valid QC level

  assign(allfinalwb, copyWorkbook(wb))

    #Create a sheet for the SN
    addWorksheet(get(allfinalwb), usn[i])

    #Write QC assay values to workbook object and create worksheets for every instrument?
    writeData(get(allfinalwb), usn[i], get(allfinalqc),
              colNames = TRUE, startRow = assaystart)

    #Use WriteData to write QC data to already created sheet
    writeData(get(allfinalwb), usn[i], get(allfinaldata),
              startRow = nrow(get(allfinalqc))+assaystart+datastart/2, colNames = TRUE)

    #Add 1 extra digits for  WBC LYM %L MID %M GRAN %G HGB HCT MCV MCH MCHC
    addStyle(get(allfinalwb), usn[i], onedigit,
             rows = assaystart:c(nrow(get(allfinaldata))+nrow(get(allfinalqc))+assaystart+datastart),
             cols = grep("WBC|LYM|%L|MID|%M|GRAN|%G|HGB|HCT|MCV|MCH|MCHC", finalcols),
             gridExpand = TRUE, stack = TRUE)

    #Add 2 extra digits for RBC
    addStyle(get(allfinalwb), usn[i], twodigit,
             rows = assaystart:c(nrow(get(allfinaldata))+nrow(get(allfinalqc))+assaystart+datastart),
             cols = grep("RBC", finalcols), gridExpand = TRUE, stack = TRUE)

    #Add centered formatting for finalcols, disable if it works through formattable
    addStyle(get(allfinalwb), usn[i], centered,
             rows = assaystart:c(nrow(get(allfinaldata))+nrow(get(allfinalqc))+assaystart+datastart),
             cols = 1:length(finalcols), gridExpand = TRUE, stack = TRUE)

    #Add bold assay value colnames
    addStyle(get(allfinalwb), usn[i], bold,
             rows = c(assaystart, alldatastart),
             cols = 1:length(finalcols), gridExpand = TRUE, stack = TRUE)

    #Add permanent red formatting in the future withto cells that are OOR starting in row 7
    #addStyle(get(finalwb), usn[i], red, rows = x, cols = y, gridExpand = TRUE, stack = TRUE)

    #Use conditional formatting for now, could test "red" addstyle but probably too slow
    
    #for loop to add to each specific lot to the wb
    for (k in 1:length(valuecols)) {
      
      
      
      #reset start position of conditional formatting for each new LOT
      s <- alldatastart+1
      
    #for loop to add to each column of that section
     for (j in 1:length(ulevel)) {
     
       #assign search "level sn" change lot to L N H for each element
       lotdata <- paste("LNH", usn[i], "_" , ulevel[j], sep = "")
       
       #check if sn/qc combo exists in list.qc
       if (!exists(lotdata)) {next}
       
       #less than rule
       conditionalFormatting(get(allfinalwb), usn[i], cols = k+1,
                             rows = s:c(nrow(get(lotdata))+s-1),
                             usn[i], style = red, type = "expression",                    #c(lower limit, upper limit)
                             rule = paste("<", c(get(paste("QCLNH", qclevel[j], sep = ""))[1,valuecols[k]] - get(paste("QCLNH", qclevel[j], sep = ""))[2,valuecols[k]])))
       #greater than rule
       conditionalFormatting(get(allfinalwb), usn[i], cols = k+1,
                             rows = s:c(nrow(get(lotdata))+s-1),
                             usn[i], style = red, type = "expression",                    #c(lower limit, upper limit)
                             rule = paste(">", c(get(paste("QCLNH", qclevel[j], sep = ""))[1,valuecols[k]] + get(paste("QCLNH", qclevel[j], sep = ""))[2,valuecols[k]])))
       #Add to row counter start position
       s <- s + nrow(get(lotdata))
       #add a +1 here ^?
     } #end k loop
      
      #row counter for plots
      r = 2
     } #end j loop ulevel
     
      

    #switch to turn graphs on or off:
    if (graphs) {
    
    #Add timeseries ggplots of qc data one at a time for every variable, make plot then insertplot into wb object.
    #use r for row # to space out graphs
    for (m in 1:length(valuecols)) {

      #Comments df for plotting, remove any NA rows
      events <- remove_missing(data.frame(get(allfinaldata)["Comments"], get(allfinaldata)[paste("datetime", str_sub(valuecols[m], start = -2, end = -1), sep = "_")]), na.rm = TRUE)
      
      #get datetime x-axis, choose LL, NL, or HL
      plot <- ggplot(data = get(allfinaldata), aes_string(x= paste("datetime", str_sub(valuecols[m], start = -2, end = -1), sep = "_"), y = valuecols[m])) + geom_line(color = "blue", na.rm = TRUE)

      #Add title to graph
      plot <- plot + ggtitle(paste("Emerald 18 SN", usn[i], valuecols[m])) + theme_bw()

      #Modify scale
      #plot <- plot + scale_x_datetime(date_breaks = "2 month", date_minor_breaks = "1 week", date_labels = "%D")
      
      #Add Target
      #plot <- plot + geom_hline(yintercept =  get(paste("QCLNH", qclevel[i], sep = ""))[1,valuecols[m]],
      #                          "line", color = "red")

      #Add upper limit
      #plot <- plot + geom_hline(yintercept =  get(paste("QCLNH", qclevel[i], sep = ""))[1,valuecols[m]] + get(paste("QCLNH", qclevel[i], sep = ""))[2,valuecols[m]],
      #                          "line", color = "orange")

      #Add lower limit
      #plot <- plot + geom_hline(yintercept =  get(paste("QCLNH", qclevel[i], sep = ""))[1,valuecols[m]] - get(paste("QCLNH", qclevel[i], sep = ""))[2,valuecols[m]],
      #                          "line", color = "orange")
      
      #Add Comment vertical lines
      plot <- plot + geom_vline(data = events, aes_string( xintercept = paste("datetime", str_sub(valuecols[m], start = -2, end = -1), sep = "_")))
      
      #Add text to vertical lines
      #plot <-  plot + geom_text(data = events, aes_string( label = Comments,))
      
      
      #to show plot in viewer for insertplot function
      print(plot)

      insertPlot(get(allfinalwb), usn[i], startCol = length(finalcols)+3, startRow =  r)

      r= r + 20
    }  
    } #if statement close to turn graphs on/off                                                                                                                             #this code written by JIT


  #Write wb object to file and name after QC level
  #Use tryCatch to go to next if file exists
     tryCatch(expr = saveWorkbook(get(allfinalwb),
                                paste(folder, "/E18 All QC for SN ", usn[i], ".xlsx", sep = ""), overwrite = FALSE),
              finally = next)

} #i usn bracket close


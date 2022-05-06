#Method for cleaning genetics reports 

########## Initialization ##########

suppressPackageStartupMessages({
  library(tidyverse) #needs to be installed
  library("methods")
  library("fs")
  library("openxlsx") #needs to be installed
})





#take the argument and set it as the path to the file you're looking to parse
path_to_file <- "C:/Users/adwalker/Documents/projects/Muscle Set/Copath_reports/converting_prn/prnt0002_converted.xlsx"

#read the file and change NA to blank so that it's able to be processed

suppressMessages( #Needed to suppress the creation of new column names due to col_names = FALSE
  test_import_type_excel <- read.xlsx(path_to_file,colNames = FALSE,detectDates=TRUE)
)

test_import_type_excel[is.na(test_import_type_excel)] <- ""

processing_data <- as.data.frame(test_import_type_excel)

#Need to figure out what type of report you are parsing (either Sanford or Invitae)

file_name <- path_file(path_to_file)

########## Cleaning the dataset to make it more easily parsible ##########

#cleaning function to get rid of NA
not_all_na <- function(x) any(!is.na(x))

#Cleaning data from CSV 
compressed <- processing_data %>% select(where(not_all_na)) %>%
  .[!apply(.=="",1,all),]
#t()

#Creating duplicate file to double check against later
old_compressed_to_check_against <- compressed

if(dim(processing_data)[2] > 1){

  #Shifting values to get rid of extra empty cells --> recreating graphs
  for (row in 1:dim(compressed)[1]){
    for(column in 1:dim(compressed)[2]) {
      
      active_column = column
      
      if(identical(compressed[row,active_column],"")) {
        
        active_column = column + 1
        
        #move through the rest of the cells in the column to see if there are any more values to shove over
        while( active_column < dim(compressed)[2]+1) {
          if(!identical(compressed[row,active_column],"")) {
            compressed[row,column] <- compressed[row,active_column]
            compressed[row,active_column] <- ""
            active_column = 1000 #this breaks out of the column structure and ends the while loop
          } else {
            active_column = active_column + 1
          }
        }
      }
    }
  }
}


########## Pull out wanted chart(s) ##########

#Pull out wanted information based off which genetics report we are looking at






########## Output the File ##########

output_file_name <- paste(str_sub(file_name,end = -6), "_output_chart.csv", collapse = "",sep = "")



write.csv(as.data.frame(compressed),output_file_name,row.names = FALSE)
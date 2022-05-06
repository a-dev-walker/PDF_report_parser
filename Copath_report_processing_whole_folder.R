#Method for cleaning integrated Copath reports
#Takes the input of a folder selected using the directory

########## Initialization ##########

suppressPackageStartupMessages({
  library("tidyverse") #needs to be installed
  library("methods")
  library("fs")
  library("openxlsx") #needs to be installed
})


########## Data Import ##########


#making sure that there is only one argument supplied to the program
main_dir <- choose.dir(caption = "Please select folder containing copath reports to be processed from excel:")
main_dir <- gsub("\\\\", "/", main_dir)




#Making sure to only take .xlsx files
file_list <- list.files(path=main_dir, pattern = "*.xlsx",full.names = TRUE)




#Function for processing individual files and returning their desired data and file path for creating output file
Process_individual_file <- function(copath_file_path) {
  
  #take the argument and set it as the path to the file you're looking to parse
  path_to_file <- copath_file_path[1]
  
  #Read the file and change NA to blank so that it's able to be processed
  #Needed to suppress the creation of new column names due to col_names = FALSE
  suppressMessages( 
    test_import_type_excel <- read.xlsx(path_to_file,colNames = FALSE,detectDates=TRUE)
  )
  
  test_import_type_excel[is.na(test_import_type_excel)] <- ""
  
  processing_data <- as.data.frame(test_import_type_excel)
  
  
  
  file_name <- path_file(path_to_file)
  
  ########## Cleaning the dataset to make it more easily parsable ##########
  
  
  
  #Cleaning function to get rid of NA
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
          
          #Move through the rest of the cells in the column to see if there are any more values to shove over
          while( active_column < dim(compressed)[2]+1) {
            if(!identical(compressed[row,active_column],"")) {
              compressed[row,column] <- compressed[row,active_column]
              compressed[row,active_column] <- ""
              active_column = 1000 #this breaks out of the column structure and ends the while loop --> It's prohibitively large so as to not be achieved under non-abnormal situations
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
  
  desired_data <- compressed
  
  
  
  
  
  
  
  ########## Output the File ##########
  
  output_file_name <- paste(str_sub(file_name,end = -6), "_output_chart.csv", collapse = "",sep = "")
  
  return(list(desired_data,output_file_name))
} #End of Process_individual_file function




#Creating a sub directory to put the output files in 
setwd(main_dir)
sub_dir <- "processing_outputs"




#Processing the data and exporting it to newly created csv files in the new sub-directory
for(file_for_converting in file_list){
  
  file_for_converting <- gsub("[~$]", "",file_for_converting)
  
  
  processed_data <- Process_individual_file(file_for_converting)
  
  output_file_data <- as.data.frame(processed_data[1])
  output_file_name <- toString(processed_data[2])
  
  
  
  if (file.exists(sub_dir)){
    setwd(file.path(main_dir, sub_dir))
  } else {
    dir.create(file.path(main_dir, sub_dir))
    setwd(file.path(main_dir, sub_dir))
  }
  
  #write_csv(output_file_data,output_file_name,col_names=FALSE)
  write.csv(output_file_data,output_file_name,row.names=FALSE,fileEncoding = "Windows-1252")
  
  setwd(main_dir)
  
}
## Functions to convert an Metadatasheet to a SumExp

# Source the script! (multiple funcitons in here)

# Input: 
# - completeted Metadatasheet (xlsx file)
# - flag_dataMatrix : TRUE/FALSE
#       If FALSE, transcriptomic FASTQ files are expected and processed (TOD=)

# Note that the position of the Metadatasheet within the filesystem is considered
# as root, searching for specified data further down
sheet = "../Input/current_major/Supplementary_Showcase/Mass_RNASeq/MetaDataSheet_Developmental programming of Kupffer cells by maternal obesity causes fatty liver disease in offspring_1eb2ec14-eda6-4e59-b8d8-a2edb217e676.xlsx"

library(readxl)
library(dplyr)
library(SummarizedExperiment)

MetadatasheetToSumExp <- function(sheet, flag_dataMatrix = T){

  Metadatasheet_raw <- as.data.frame(read_excel(sheet,sheet="Input"))
  # check for Measurement type:
  check_pass=T
  RowMeasurement <- which(my_data_tmp[,1]%in%"measurement type")
  measureType = Metadatasheet_raw[RowMeasurement,2]
  print(paste0("The measurement type is: ", measureType))
  if(measureType %in% c("bulk_lipidomics","bulk_metabolomics","bulk_RNA_seq")){
    check_pass = T
  }else{
    check_pass = F
    print("It should be on of: bulk_lipidomics, bulk_metabolomics or bulk_RNA_seq")
    # add that we can stop here
  }

  # Read in measurement-matching section
  SampleSection <- which(Metadatasheet_raw[,1]%in%"Sample-Section")+1 # must be present
  
  subsample_present <- Metadatasheet_raw[which(Metadatasheet_raw[,1]%in%"subsample_present"),2] == 1
  subsubsample_present <- Metadatasheet_raw[which(Metadatasheet_raw[,1]%in%"subsubsample_present"),2] == 1
  
  all_tables <- as.data.frame(read_excel(sheet,sheet="Input",skip = SampleSection,trim_ws = T,col_names = F))
  
  sample_table <- all_tables[1:which(all_tables[,1]%in%"Sub-Sample Section")-1,]
  
  sample_table_t <- as.data.frame(t(sample_table))
  colnames(sample_table_t) <- sample_table_t[1,]
  sample_table_t <- sample_table_t[-1,]
  sample_table_t <- sample_table_t[,!is.na(colnames(sample_table_t))]
  sample_table_t <- sample_table_t[rowSums(is.na(sample_table_t)) < ncol(sample_table_t),]

  merged_table <- sample_table_t
  if(subsample_present | subsubsample_present){
    if(subsample_present){
      subample_table <- all_tables[which(all_tables[,1]%in%"Sub-Sample Section")+1:nrow(all_tables),]
    }
    if(subsubsample_present){
      subsubsample_table <- subample_table[which(subample_table[,1]%in%"Sub-Sub-Sample Section")+1:nrow(subample_table),]
      subsubsample_table_t <- as.data.frame(t(subsubsample_table))
      colnames(subsubsample_table_t) <- paste0(subsubsample_table_t[1,],"_s2")
      subsubsample_table_t <- subsubsample_table_t[,!is.na(colnames(subsubsample_table_t))]
      subsubsample_table_t <- subsubsample_table_t[rowSums(is.na(subsubsample_table_t)) < ncol(subsubsample_table_t),]
      
    }
    subample_table <- subample_table[1:which(subample_table[,1]%in%"Sub-Sub-Sample Section")-1,]
    subample_table_t <- as.data.frame(t(subample_table))
    colnames(subample_table_t) <- paste0(subample_table_t[1,],"_s1")
    subample_table_t <- subample_table_t[-1,]
    subample_table_t <- subample_table_t[,!is.na(colnames(subample_table_t))]
    subample_table_t <- subample_table_t[rowSums(is.na(subample_table_t)) < ncol(subample_table_t),]
    
    if(subsubsample_present){
      merged_table <- subample_table_t %>%
        left_join(subsubsample_table_t, by = c("global_ID" = "sub_sample_match_s2")) 
      merged_table <- sample_table_t %>%
        left_join(merged_table, by = c("global_ID" = "sample_match_s1")) 
    }else{
      merged_table <- sample_table_t %>%
        left_join(subample_table_t, by = c("global_ID" = "sample_match_s1")) 
    }
  }
  
  # the last level is measurement level hence here the measuremen IDs expected
  # Identify the latest personal_ID to put as rownmaes
  rownames(merged_table) <- merged_table[,tail(which(grepl(c("personal_ID|personal_ID_s1|personal_ID_2"),colnames(merged_table))),n=1)]

  # Identify data to add
  # Check out data file linkage
  
  DataFiles_position <- which(Metadatasheet_raw[,1]%in%"DataFiles-Linkage")
  dataFilesLinkgae <- which(Metadatasheet_raw[,2] %in% "Link Type Personal ID to provided data")
  dataFilesLinkgae_options <- Metadatasheet_raw[tail(dataFilesLinkgae,n=1),3]
  if(flag_dataMatrix){
    #data matrix given under processed data
    filename_to_search <- Metadatasheet_raw[tail(dataFilesLinkgae,n=1)+2,3]
    print(paste0("The Data Files option is: ",tail(dataFilesLinkgae_options,n=1)))
    # if comment is empty return to user all filenames in reach from root Metadatasheet to choose from
    if(is.na(filename_to_search) | !grepl("\\.xlsx$|\\.csv$|\\.txt$", filename_to_search, ignore.case = TRUE)){
      filename_to_search <- choose_file(dirname(sheet))

    }
    dataFile <- read_file_into_dataframe(file.path(dirname(sheet),filename_to_search))
    # check if first column holds non-numeric stuff then put to rownames
    dataMatrix <- check_and_convert_first_column(dataFile)
    
    # check rownames for IDs given in Metadatasheet
    # As R does not like numbers as column names check if given id is a subset of current column names
    # if so give user a chance to see what is matched and then to agree or not
    matchResults <- match_rownames_to_columns(merged_table, dataMatrix)
    # most likely easier to change Metadata IDs hence  dataMatrix IDs are taken as final IDs
    exact_matches <- matchResults[,2]
    rownames(merged_table) <- ifelse(rownames(merged_table) %in% exact_matches, rownames(merged_table), exact_matches)
    dataMatrix <- dataMatrix[,rownames(merged_table)]
    
    entitie = data.frame(row.names=rownames(dataMatrix),name=rownames(dataMatrix))
    
    # create additional Metadata (can be advanced!)
    Metadatasheet_raw[1:which(Metadatasheet_raw[,1]%in%"Sample-Section"),1:4]
    segmentInfo <- c("General","Experimental System","covariates / constants","Time-Dependence-timeline","Preparation","Measurement")
    extracted_info_list <- list()
    
    # Iterate through each item
    for (item in segmentInfo) {
      # Find rows where the item is present in the first column
      item_rows <- which(Metadatasheet_raw[,1] == item)[1]
      # Extract information from the found positions till the next complete empty row
      next_empty_row <- names(which(rowSums(!is.na(Metadatasheet_raw[item_rows:nrow(Metadatasheet_raw),])) == 0)[1])

      # Extract information
      extracted_info <- Metadatasheet_raw[item_rows:next_empty_row, ]
      next_empty_col <- min(which(colSums(!is.na(extracted_info)) == 0))
      
      
      # Store the extracted information in the list with the item name
      extracted_info_list[[item]] <- as.data.frame(extracted_info[-1,1:next_empty_col])
    }
    
    ## Lets Make a SummarizedExperiment Object for reproducibility and further usage
    Metadata_SumExp=
      SummarizedExperiment(assays  = dataMatrix,
                           colData = merged_table,
                           rowData = entitie,
                           metadata = extracted_info_list
      )
    
    return(Metadata_SumExp)
  }
}


# Read in Excel File
# Search for sample_section save row number
# read in again and skip first row number lines
fun_readInSampleTable <- function(dataFileName){
  my_data_tmp <- as.data.frame(read_excel(dataFileName,sheet="Input"))
  RowsToSkip <- which(my_data_tmp[,1]%in%"Sample-Section")+1
  my_data_tmp <- as.data.frame(read_excel(dataFileName,sheet="Input",skip = RowsToSkip))
  
  # Advance: check if subsample etc are present 
  # for now remove any non complete rows
  my_data_tmp <- my_data_tmp[!is.na(my_data_tmp$personal_ID),]
  
  my_data_tmp <- my_data_tmp[-(which(my_data_tmp[,1] %in% "Sub-Sample Section"):nrow(my_data_tmp)),]
  
  my_data_tmp <- t(my_data_tmp)
  colnames(my_data_tmp) <- my_data_tmp[1,]
  my_data_tmp <- as.data.frame(my_data_tmp[-1,])
  
  return(my_data_tmp) 
}

# TODO: For Future somehow implempent type check, eg if there are numerics!!
# either by prior info (e.g. weight, days, BMI is numeric) or by going through cols and check wheter to numeric is possible?!


choose_file <- function(path=".") {
  # Get the list of files in the current working directory
  files <- list.files(path=path, recursive = T)
  
  # Filter files based on suffix
  valid_files <- files[grep("\\.xlsx$|\\.csv$|\\.txt$", files, ignore.case = TRUE)]
  
  # Check if multiple files are found
  if (length(valid_files) == 0) {
    cat("No valid files found in the current directory.\n")
    return(NULL)
  } else if (length(valid_files) == 1) {
    cat("Found the following file:\n")
    selected_file <- valid_files[1]
  } else {
    cat("Multiple valid files found in the current directory:\n")
    
    # Prompt user to choose a file
    choice <- utils::menu(title = "File Selection", choices = valid_files, graphics = F)
    if (choice == 0) {
      cat("User canceled file selection.\n")
      return(NULL)
    }
    
    selected_file <- valid_files[choice]
  }
  
  cat("Selected file:", selected_file, "\n")
  return(selected_file)
}

read_file_into_dataframe <- function(filename) {
  # Determine the file suffix
  suffix <- tools::file_ext(filename)
  
  # Read the file into a data frame based on the suffix
  if(tolower(suffix) %in% c("xlsx", "xls")){
    data <- readxl::read_excel(filename)
  }else if (tolower(suffix) == "csv"){
    data <- read.csv(filename)
  }else if (tolower(suffix) == "txt"){
    data <- read.table(filename, header = TRUE, sep = "\t")
  }else{
    stop("Unsupported file format. Please provide a file with 'xlsx', 'csv', or 'txt' extension.")
  }
  
  return(as.data.frame(data))
}
    
check_and_convert_first_column <- function(data) {
  # Check if data is a data frame
  if (!is.data.frame(data)) {
    stop("Input is not a data frame.")
  }
  
  # Check and convert the first column
  first_col <- data[, 1]
  converted_col <- suppressWarnings(as.numeric(first_col))
  
  # Check if any non-numeric characters are present
  if (any(is.na(converted_col))) {
    # Some non-numeric characters are present
    rownames(data) <- as.character(data[, 1])

    
    # Remove the first column from the data frame
    res <- data[, -1]
    
    cat("Non-numeric characters found in the first column. Converted strings to row names.\n")
  } else {
    # All values are numeric
    cat("All values in the first column are numeric.\n")
  }
  
  return(res)
}


match_rownames_to_columns <- function(merged_table, dataMatrix) {
  # Check if merged_table has row names
  if (!is.data.frame(merged_table) || is.null(rownames(merged_table))) {
    stop("Input merged_table is not a data frame with row names.")
  }
  
  # Check if dataMatrix has column names
  if (is.null(colnames(dataMatrix))) {
    stop("Input dataMatrix is not a matrix with column names.")
  }
  
  # Initialize a data frame to store the matching results
  matching_results <- data.frame(RowName = character(0), MatchedColumn = character(0), stringsAsFactors = FALSE)
  
  # Iterate through row names of merged_table
  for (row_name in rownames(merged_table)) {
    # Check for an exact match
    exact_match <- row_name %in% colnames(dataMatrix)
    
    if (!exact_match) {
      # Check for a subset match
      subset_matches <- grep(row_name, colnames(dataMatrix), value = TRUE, ignore.case = TRUE)
      
      # If there are subset matches, use fix to allow user inspection
      if (length(subset_matches) > 0) {
        cat("Row name:", row_name, "\n")
        cat("Subset matches found:", subset_matches, "\n")
        
        # Use fix to let the user inspect and potentially correct the matching
        fixed_matches <- data.frame(RowName = row_name, SubsetMatches = subset_matches)
        if (!is.null(fixed_matches)) {
          matching_results <- rbind(matching_results, fixed_matches)
        }
      } else {
        # No matches found
        cat("No match found for row name:", row_name, "\n")
      }
    } else {
      # Exact match found
      matching_results <- rbind(matching_results, data.frame(RowName = row_name, MatchedColumn = row_name))
    }
    
  }

  return(matching_results)
}





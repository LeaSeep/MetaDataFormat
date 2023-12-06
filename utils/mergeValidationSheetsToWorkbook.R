# mergeValidation.R

library(openxlsx)

mergeValidationSheets <- function(workbookList, outputWorkbook,sheet2extract) {
  # Create a new workbook for merged validation sheets
  mergedWorkbook <- openxlsx::createWorkbook()
  
  # Loop through each workbook in the list
  for (i in seq_along(workbookList)) {
    # Load the workbook
    currentWorkbook <- openxlsx::readWorkbook(workbookList[i],sheet = sheet2extract)

    # Add the "Validation" sheet to the merged workbook

    # If it's the first workbook, add the sheet as the very first sheet
    sheet_name = paste0(i,"_",sheet2extract)
    addWorksheet(mergedWorkbook, sheet_name)
    writeData(mergedWorkbook, sheet_name, currentWorkbook)

  # Save the merged workbook
  saveWorkbook(mergedWorkbook, outputWorkbook,overwrite = T)
  }
}

# Provide a list of workbook filenames
workbookList <-list.files(path="../Input/current_major",pattern="^MetaDaten_.*\\.xlsm$",recursive = T,full.names = T)
 
workbookList <-workbookList[c(21,1:20,22,23)]

# Specify the output workbook name
outputWorkbook <- "Dependent_v1.9.xlsx"

# Run the function with the specified arguments
mergeValidationSheets(workbookList, outputWorkbook,sheet2extract="dependentFields")





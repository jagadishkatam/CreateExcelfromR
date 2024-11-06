library(openxlsx)
library(dplyr)
library(purrr)

# Create a new workbook
wb <- createWorkbook()

# Create a summary worksheet
summary_sheet <- "Cover Page"
addWorksheet(wb, sheetName = summary_sheet)


iris$Species <- as.character(iris$Species)
str(iris)

# Define the categories to split by (unique species in this case)
cats <- iris %>% select(Species) %>% distinct() %>% pull(Species)

typeof(cats)

# Function to add data for each species to its own worksheet
workbook <- function(species) {
  # Add a worksheet named after the species
  addWorksheet(wb, sheetName = species)
  
  # Create a header style
  header_style <- createStyle(halign = "center", textDecoration = "bold")
  
  # Filter data for the specific species
  data_subset <- iris %>% filter(Species == species)
  
  # Write the filtered data to the worksheet
  writeData(wb, sheet = species, x = data_subset, headerStyle = header_style)
}

# Apply the function to create a worksheet for each species
map(cats, ~workbook(.x))

# Write the sheet names in the summary worksheet and add hyperlinks
writeData(wb, sheet = summary_sheet, x = data.frame(Sheet_Names = cats), startCol = 1, startRow = 1)

# Add hyperlinks to each sheet
for (i in seq_along(cats)) {
  sheet_name <- cats[i]
  writeFormula(wb, summary_sheet, startCol = 2, startRow = i + 1, 
               x =  paste0('=HYPERLINK("',"#'",'" &A', i + 1 , '& "',"'",'!A1",A',i + 1,')') 
  )
}

# Save the workbook
saveWorkbook(wb, file = "./data/sample_with_hyperlinks.xlsx", overwrite = TRUE)

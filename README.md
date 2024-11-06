# CreateExcelfromR

# Iris Dataset Excel Export with Hyperlinked Worksheets

This R script generates an Excel workbook using the `iris` dataset, with each unique species categorized into individual worksheets. Additionally, a cover page with hyperlinks is created to easily navigate between the species-specific sheets.

## Features

- **Workbook Creation**: Initializes a new workbook and adds a summary worksheet.
- **Sheet Generation**: Creates individual worksheets for each species in the `iris` dataset.
- **Data Filtering**: Filters the `iris` data by species and writes the filtered data to each corresponding worksheet.
- **Hyperlinking**: Generates a cover page with hyperlinks to each species-specific sheet for quick access.

## Requirements

- R (version X.X or later)
- Packages: 
  - `openxlsx`: For creating and manipulating Excel files.
  - `dplyr`: For data manipulation.
  - `purrr`: For functional programming to map operations across data.

## Installation

Install the required packages if they are not already available:

```r
install.packages("openxlsx")
install.packages("dplyr")
install.packages("purrr")
```

<!-- This file is used to create README.md
-->

# importExcel

importExcel is a department package to import data that has been checked
with the internal UHN Biostatistics `dataChecker`.

The package has one main function which will read in the excel data, add
the labels and recode factors based on the information provided in the
data dictionary.

## Installation

You can install the latest version of importExcel from
[GitHub](https://github.com/) with:

``` r
# install.packages("devtools")
devtools::install_github("biostatsPMH/importExcel")
```

## Documentation

[Online Documentation](https://biostatsPMH.github.io/importExcel/)

## Example 1: Tidy Data

This is a nice clean data set

``` r
library(importExcel)
data_file <- system.file("extdata", "study_data.xlsx", package = "importExcel")

# This will read in the file, add variable labels and recode variables to 
# factors as specified in the dictionary sheet
file_import <- read_excel_with_dictionary(data_file = data_file,
                                          data_sheet = "data",dictionary_sheet = "dict")

# extract the recoded-data
coded_data <- file_import$coded_data

# file_import contains the following objects:
# # check warnings
# file_import$warnings
# 
# # check for text omitted from numeric data
# file_import$numeric_converstions
# 
# # check for attempted date conversions
# file_import$date_converstions
# 
# # Look at the dictionary used
# file_import$updated_dictionary
```

Describe the coded data:

``` r
require(tidyverse)
require(reportRmd)

coded_data |> 
  select(!studyID) |> 
  rm_compactsum(xvars=everything())
#> no statistical tests will be applied to date variables, date variables will be summarised with median
```

<table class="table table" style="margin-left: auto; margin-right: auto; margin-left: auto; margin-right: auto;">
<thead>
<tr>
<th style="text-align:left;position: sticky; top:0; background-color: #FFFFFF;">
</th>
<th style="text-align:right;position: sticky; top:0; background-color: #FFFFFF;">
Full Sample (n=100)
</th>
</tr>
</thead>
<tbody>
<tr>
<td style="text-align:left;">
<span style="font-weight: bold;">Age at diagnosis</span>
</td>
<td style="text-align:right;">
44.0 (32.0-57.2)
</td>
</tr>
<tr>
<td style="text-align:left;">
<span style="font-weight: bold;">sex - Male</span>
</td>
<td style="text-align:right;">
45 (45%)
</td>
</tr>
<tr>
<td style="text-align:left;">
<span style="font-weight: bold;">BMI</span>
</td>
<td style="text-align:right;">
29.0 (23.4-34.2)
</td>
</tr>
<tr>
<td style="text-align:left;">
<span style="font-weight: bold;">heart rate (bmp)</span>
</td>
<td style="text-align:right;">
84.0 (72.0-92.0)
</td>
</tr>
<tr>
<td style="text-align:left;">
<span style="font-weight: bold;">date of diagnosis</span>
</td>
<td style="text-align:right;">
2021-11-12 (2020-11-18 to 2023-06-18)
</td>
</tr>
<tr>
<td style="text-align:left;">
<span style="font-weight: bold;">date of randomisation</span>
</td>
<td style="text-align:right;">
2023-06-25 (2023-04-18 to 2023-08-30)
</td>
</tr>
<tr>
<td style="text-align:left;">
<span style="font-weight: bold;">ECOG</span>
</td>
<td style="text-align:right;">
1.0 (0.0-3.0)
</td>
</tr>
<tr>
<td style="text-align:left;">
<span style="font-weight: bold;">Treatment</span>
</td>
<td style="text-align:right;">
47 (47%)
</td>
</tr>
<tr>
<td style="text-align:left;">
<span style="font-weight: bold;">baseline score</span>
</td>
<td style="text-align:right;">
61.5 (27.5-76.2)
</td>
</tr>
<tr>
<td style="text-align:left;">
<span style="font-weight: bold;">final score</span>
</td>
<td style="text-align:right;">
42.5 (18.5-81.5)
</td>
</tr>
<tr>
<td style="text-align:left;">
<span style="font-weight: bold;">hospital</span>
</td>
<td style="text-align:right;">
</td>
</tr>
<tr>
<td style="text-align:left;padding-left: 2em;" indentlevel="1">
PMH
</td>
<td style="text-align:right;">
35 (35%)
</td>
</tr>
<tr>
<td style="text-align:left;padding-left: 2em;" indentlevel="1">
Western
</td>
<td style="text-align:right;">
33 (33%)
</td>
</tr>
<tr>
<td style="text-align:left;padding-left: 2em;" indentlevel="1">
Krembil
</td>
<td style="text-align:right;">
32 (32%)
</td>
</tr>
<tr>
<td style="text-align:left;">
<span style="font-weight: bold;">Diagnosis</span>
</td>
<td style="text-align:right;">
</td>
</tr>
<tr>
<td style="text-align:left;padding-left: 2em;" indentlevel="1">
lupus
</td>
<td style="text-align:right;">
24 (24%)
</td>
</tr>
<tr>
<td style="text-align:left;padding-left: 2em;" indentlevel="1">
mumps
</td>
<td style="text-align:right;">
24 (24%)
</td>
</tr>
<tr>
<td style="text-align:left;padding-left: 2em;" indentlevel="1">
measles
</td>
<td style="text-align:right;">
31 (31%)
</td>
</tr>
<tr>
<td style="text-align:left;padding-left: 2em;" indentlevel="1">
rubella
</td>
<td style="text-align:right;">
21 (21%)
</td>
</tr>
</tbody>
</table>

## Example 2: Messier Data

This data has messy dates which are converted to yyyy-mm-dd format for
consistency

``` r

data_file <- system.file("extdata", "study_data_messy.xlsx", package = "importExcel")

# This will read in the file, add variable labels and recode variables to 
# factors as specified in the dictionary sheet
file_import <- read_excel_with_dictionary(data_file = data_file,
                                          data_sheet = "data",dictionary_sheet = "dict")
#> New names:
#> • `diagnosis` -> `diagnosis...13`
#> • `diagnosis` -> `diagnosis...14`

# extract the recoded-data
coded_data <- file_import$coded_data
```

Describe the coded data:

``` r
coded_data |> 
  select(!studyID) |> 
  rm_compactsum(xvars=everything())
#> no statistical tests will be applied to date variables, date variables will be summarised with median
```

<table class="table table" style="margin-left: auto; margin-right: auto; margin-left: auto; margin-right: auto;">
<thead>
<tr>
<th style="text-align:left;position: sticky; top:0; background-color: #FFFFFF;">
</th>
<th style="text-align:right;position: sticky; top:0; background-color: #FFFFFF;">
Full Sample (n=100)
</th>
<th style="text-align:right;position: sticky; top:0; background-color: #FFFFFF;">
Missing
</th>
</tr>
</thead>
<tbody>
<tr>
<td style="text-align:left;">
<span style="font-weight: bold;">Age at Diagnosis</span>
</td>
<td style="text-align:right;">
44.0 (32.0-57.2)
</td>
<td style="text-align:right;">
0
</td>
</tr>
<tr>
<td style="text-align:left;">
<span style="font-weight: bold;">Sex</span>
</td>
<td style="text-align:right;">
2.0 (1.0-2.0)
</td>
<td style="text-align:right;">
0
</td>
</tr>
<tr>
<td style="text-align:left;">
<span style="font-weight: bold;">BMI</span>
</td>
<td style="text-align:right;">
29.0 (23.4-34.2)
</td>
<td style="text-align:right;">
0
</td>
</tr>
<tr>
<td style="text-align:left;">
<span style="font-weight: bold;">Heart Rate</span>
</td>
<td style="text-align:right;">
84.0 (72.0-92.0)
</td>
<td style="text-align:right;">
0
</td>
</tr>
<tr>
<td style="text-align:left;">
<span style="font-weight: bold;">date of diagnosis</span>
</td>
<td style="text-align:right;">
2021-11-21 (2020-12-26 to 2023-06-18)
</td>
<td style="text-align:right;">
0
</td>
</tr>
<tr>
<td style="text-align:left;">
<span style="font-weight: bold;">date of randomisation</span>
</td>
<td style="text-align:right;">
2023-06-21 (2023-04-03 to 2023-08-27)
</td>
<td style="text-align:right;">
0
</td>
</tr>
<tr>
<td style="text-align:left;">
<span style="font-weight: bold;">ECOG</span>
</td>
<td style="text-align:right;">
1.0 (0.0-3.0)
</td>
<td style="text-align:right;">
0
</td>
</tr>
<tr>
<td style="text-align:left;">
<span style="font-weight: bold;">Treatment</span>
</td>
<td style="text-align:right;">
47 (47%)
</td>
<td style="text-align:right;">
0
</td>
</tr>
<tr>
<td style="text-align:left;">
<span style="font-weight: bold;">baseline score</span>
</td>
<td style="text-align:right;">
61.5 (27.5-76.2)
</td>
<td style="text-align:right;">
0
</td>
</tr>
<tr>
<td style="text-align:left;">
<span style="font-weight: bold;">final score</span>
</td>
<td style="text-align:right;">
42.5 (17.5-82.5)
</td>
<td style="text-align:right;">
2
</td>
</tr>
<tr>
<td style="text-align:left;">
<span style="font-weight: bold;">Hospital</span>
</td>
<td style="text-align:right;">
</td>
<td style="text-align:right;">
0
</td>
</tr>
<tr>
<td style="text-align:left;padding-left: 2em;" indentlevel="1">
PMH
</td>
<td style="text-align:right;">
35 (35%)
</td>
<td style="text-align:right;">
</td>
</tr>
<tr>
<td style="text-align:left;padding-left: 2em;" indentlevel="1">
Western
</td>
<td style="text-align:right;">
33 (33%)
</td>
<td style="text-align:right;">
</td>
</tr>
<tr>
<td style="text-align:left;padding-left: 2em;" indentlevel="1">
Krembil
</td>
<td style="text-align:right;">
32 (32%)
</td>
<td style="text-align:right;">
</td>
</tr>
<tr>
<td style="text-align:left;">
<span style="font-weight: bold;">diagnosis</span>
</td>
<td style="text-align:right;">
</td>
<td style="text-align:right;">
0
</td>
</tr>
<tr>
<td style="text-align:left;padding-left: 2em;" indentlevel="1">
measles
</td>
<td style="text-align:right;">
24 (24%)
</td>
<td style="text-align:right;">
</td>
</tr>
<tr>
<td style="text-align:left;padding-left: 2em;" indentlevel="1">
mumps
</td>
<td style="text-align:right;">
24 (24%)
</td>
<td style="text-align:right;">
</td>
</tr>
<tr>
<td style="text-align:left;padding-left: 2em;" indentlevel="1">
rubella
</td>
<td style="text-align:right;">
31 (31%)
</td>
<td style="text-align:right;">
</td>
</tr>
<tr>
<td style="text-align:left;padding-left: 2em;" indentlevel="1">
chicken pox
</td>
<td style="text-align:right;">
21 (21%)
</td>
<td style="text-align:right;">
</td>
</tr>
<tr>
<td style="text-align:left;">
<span style="font-weight: bold;">diagnosis2</span>
</td>
<td style="text-align:right;">
</td>
<td style="text-align:right;">
0
</td>
</tr>
<tr>
<td style="text-align:left;padding-left: 2em;" indentlevel="1">
measles/chicken pox
</td>
<td style="text-align:right;">
48 (48%)
</td>
<td style="text-align:right;">
</td>
</tr>
<tr>
<td style="text-align:left;padding-left: 2em;" indentlevel="1">
mumps
</td>
<td style="text-align:right;">
30 (30%)
</td>
<td style="text-align:right;">
</td>
</tr>
<tr>
<td style="text-align:left;padding-left: 2em;" indentlevel="1">
rubella
</td>
<td style="text-align:right;">
22 (22%)
</td>
<td style="text-align:right;">
</td>
</tr>
</tbody>
</table>

Look at the numeric values not imported:

``` r
file_import$numeric_conversions
#> $final_score
#> [1] "non-numeric values removed: NA, not recorded"
```

Look at the converted date values (not run):

``` r
file_import$date_conversions
```

## Sample Data

The package installs two sample files: `study_data.xlsx` and
`study_data_messy.xlsx`

The following code can be copied and run to view the first file:

``` r
tidy_data <- "study_data.xlsx"
messy_data <- "study_data_messy.xlsx"

file_path <- system.file("extdata", tidy_data, package = "importExcel")

# Check if the file exists
if (file.exists(file_path)) {
  # Open the file using the default system application (Excel)
  
  # For Windows
  if (Sys.info()['sysname'] == "Windows") {
    shell.exec(file_path)
  }
  
  # For macOS
  else if (Sys.info()['sysname'] == "Darwin") {
    system(paste("open", shQuote(file_path)))
  }
  
  # For Linux
  else if (Sys.info()['sysname'] == "Linux") {
    system(paste("xdg-open", shQuote(file_path)))
  }
  
  # If the OS is not recognized
  else {
    message("Unable to open file: Unrecognized operating system")
  }
} else {
  message("File 'study_data.xlsx' not found in the package's inst/extdata directory")
}
```

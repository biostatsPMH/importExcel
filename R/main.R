
#'Create a coded and labelled data frame from an Excel File
#'
#'Creates a cleaned, labelled and coded data set from data that the excel data
#'checker has been run on
#'
#'The variable labels from the dictionary are added, factors are created from
#'coded variables and date variables are converted to character representation.
#'
#'The following columns are expected to appear in the dictionary:
#'Current_Variable_Name,	Suggested_Name,	Label_For_Report,	Type,	Value,
#'Value_Label, Column_Number,	Import (although Suggested_Name may be missing if
#'the variable names were already clean)
#'
#'@param data_file The path to the Excel data to be imported
#'@param data_sheet The name of the sheet to import (required)
#'@param dictionary_sheet The name of the shet with the data dictionary. This
#'  must be formatted by the PMHbiostats macro file
#'@param  dates_as_character will format the dates in unambiguous yyyy-mm-dd
#'  format and return character variables. This can be useful for identifying
#'  date problems. The default is FALSE, to return date-type variables.
#'@param na Character vector of strings to interpret as missing values. This
#'  will be applied to the dictionary as well as the data
#'@importFrom readxl read_excel excel_sheets
#'@importFrom janitor excel_numeric_to_date
#'@importFrom lubridate dmy mdy ymd
#'@export
read_excel_with_dictionary <- function (data_file, data_sheet, dictionary_sheet,dates_as_character=FALSE,na="") {
  # check that the file exists
  if (!file.exists(data_file)) stop("The data file could not be found")

  # check that the sheets exist in the file
  sht_nms <- readxl::excel_sheets(data_file)
  if (!(data_sheet %in% sht_nms)) stop("The data sheet is not a sheet in the data file")
  if (!(dictionary_sheet %in% sht_nms)) stop("The dictionary sheet is not a sheet in the data file")

  dict_import <- read_excel_with_warnings(data_file, dictionary_sheet,na)
  dictionary <- dict_import$data
  # This is for older versions of the excel macro:
  last_ln_txt <- "The following variables are mostly text and should not be imported:"
  last_ln <- which(dictionary$Current_Variable_Name == last_ln_txt)
  if (length(last_ln) > 0)  dictionary <- dictionary[1:(last_ln - 1),]
  imported_data <- read_excel_with_warnings(data_file, data_sheet,na)
  if (is.null(imported_data$data))
    stop(paste("Error reading the data:\n", imported_data$errorMsg))
  new_data <- imported_data$data
  to_rm <- which(grepl("^[...]",names(new_data)))
  # columns with empty names in the data will be absent from the dictionary
  # issue a warning an proceed
  if (length(to_rm)>0){
    warning(paste("Empty column name(s) found in column(s)",paste(to_rm,collapse = ", "),". These columns will not be imported.\nTo import add a column name and re-run the dataChecker macro to update the dictionary."))
  }
  for (v in setdiff(1:ncol(new_data),to_rm)) {

    new_name <- dictionary$Suggested_Name[which(dictionary$Column_Number ==
                                                        v)]
    if (is.na(new_name)) stop(paste("The suggested name for column number",v,"is empty.\nPlease add a column name to the dictionary for this variable before import."))
    if (!(v %in% to_rm & length(new_name)==0)){
      to_rm <- setdiff(to_rm,v)
      if (substr(new_name,nchar(new_name),nchar(new_name))=="_") {
        new_name <- substr(new_name,1,nchar(new_name)-1)
        dictionary$Suggested_Name[which(dictionary$Column_Number ==
                                                v)] <- new_name
      }
      if (length(new_name) > 0) {
        if (!is.na(new_name)) {
          names(new_data)[v] <- new_name
        }
      }
    }
  }

  if ("Import" %in% names(dictionary)){
    to_rm <- c(to_rm,na.omit(dictionary[["Column_Number"]][!dictionary$Import]))
    if (length(to_rm)>0)    new_data <- new_data[,-to_rm]
    dictionary <- dictionary |>
      tidyr::fill(Import) |>
      dplyr::filter(Import)
  } else {
    if (length(to_rm)>0)    new_data <- new_data[,-to_rm]
  }

  if (any(duplicated(names(new_data)))){
    message("Duplicated variable names found.\nTo prevent this ensure distinct values in the Suggested_Name column in the data dictionary.")
    first_dupl <- which(duplicated(names(new_data)))[1]
    repeat{
      dupl_nm <- names(new_data)[first_dupl]
      dupl_ind <- which(names(new_data)==dupl_nm)
      new_names <- paste0(dupl_nm,"_",1:length(dupl_ind))
      names(new_data)[dupl_ind] <- new_names
      dictionary$Suggested_Name[which(dictionary$Column_Number %in% dupl_ind)] <- new_names
      if (!any(duplicated(names(new_data)))) break
      first_dupl <- which(duplicated(names(new_data)))[1]
    }
  }
  # Make codes into factors
  fct_variables <- dplyr::pull(dplyr::filter(dictionary,
                                             Type == "Codes"), Suggested_Name)
  for (v in fct_variables) {
    if (!(v %in% names(new_data)))
      stop(paste(v, "not found in data -check dictionary spelling and column number"))
    lvl_lbl <- dplyr::select(dplyr::filter(tidyr::fill(dictionary,
                                                       Suggested_Name, .direction = "down"), Suggested_Name ==
                                             v), Value, Value_Label)
    if (any(!is.na(lvl_lbl$Value_Label))){
      new_data[[v]] <- factor(new_data[[v]], levels = lvl_lbl$Value,
                              labels = lvl_lbl$Value_Label)
    }
  }

  # Categorical variables stay as characters
  cat_variables <- dplyr::pull(dplyr::filter(dictionary,
                                             Type == "Categorical"), Suggested_Name)
  for (v in cat_variables) {
    if (!(v %in% names(new_data)))
      stop(paste(v, "not found in data -check dictionary spelling and column number"))
    lvl_lbl <- dplyr::select(dplyr::filter(tidyr::fill(dictionary,
                                                       Suggested_Name, .direction = "down"), Suggested_Name ==
                                             v), Value, Value_Label)
    if (any(!is.na(lvl_lbl$Value_Label))){
      new_data[[v]] <- as.character(factor(new_data[[v]], levels = lvl_lbl$Value,
                                           labels = lvl_lbl$Value_Label))

    }
  }

  # Numeric variables are converted to numeric and any text is removed
  num_variables <-  dplyr::pull(dplyr::filter(dictionary,
                                              Type == "Numeric"), Suggested_Name)
  num_conversions <- NULL
  for (v in num_variables) {
    if (!(v %in% names(new_data)))
      stop(paste(v, "not found in data -check dictionary spelling and column number"))
    orig <- new_data[[v]]
    num_data <- as.numeric(orig)
    na_orig <- which(is.na(orig))
    na_new <- which(is.na(num_data))
    new_data[[v]] <- num_data
    if (length(setdiff(na_orig,na_new))>0)
      num_conversions[[v]] <- paste("non-numeric values removed:",paste(orig[setdiff(na_orig,na_new)],collapse = ", "))
  }

  date_cols <- dplyr::pull(dplyr::filter(dictionary,
                                         Type == "Date"), Suggested_Name)
  dt_conversions <- NULL
  for (v in date_cols) {
    dt_msg <- NULL
    dt_new <- sapply(new_data[[v]], function(x) {
      dt <- try(as.Date(x), silent = T)
      if (inherits(dt, "try-error"))
        dt <- try(janitor::excel_numeric_to_date(as.numeric(x)),
                  silent = T)
      if (inherits(dt, "try-error") | is.na(dt)) {
        d <- try(lubridate::dmy(trimws(x)), silent = T)
        if (is.na(d))
          d <- try(lubridate::mdy(trimws(x)), silent = T)
        if (is.na(d))
          d <- try(lubridate::ymd(trimws(x)), silent = T)
        if (!is.na(d)) {
          if (interactive())
            print(paste(x, "converted to", d))
          dt_msg <<- c(dt_msg, paste(dQuote(x), "converted to",
                                     d))
        }
        dt <- d
      }
      return(as.character(dt))
    }, USE.NAMES = F)
    dt_conversions[[v]] <- dt_msg
    if (dates_as_character) new_data[[v]] <- dt_new else new_data[[v]] <- as.Date(dt_new)
  }

  new_data <- reportRmd::set_labels(new_data, dplyr::select(dictionary,
                                                            Suggested_Name, Label_For_Report))


  return(list(coded_data = new_data,
              updated_dictionary = dictionary,
              warnings = imported_data$warnings,
              numeric_conversions = num_conversions,
              date_conversions = dt_conversions))
}

# Function to read Excel file and capture warnings

read_excel_with_warnings <- function(data_file,data_sheet,na) {
  warnings_list <- list()  # Initialize a list to store warnings
  errorMsg <- NULL
  # Use try to capture warnings
  result <- try({
    withCallingHandlers({
      # Read the Excel file
      data <- readxl::read_excel(data_file,data_sheet,na=na)
    }, warning = function(w) {
      warnings_list <<- c(warnings_list, conditionMessage(w))
      invokeRestart("muffleWarning")
    })
  }, silent = TRUE)

  # Check if there was an error
  if (inherits(result, "try-error")) {
    data <- NULL
    errorMsg <- result
    warning("An error occurred while reading the Excel file.")
  } else {
    data <- result
  }

  # Return both the data and the list of warnings
  list(data = data, warnings = warnings_list, errorMsg = errorMsg)
}

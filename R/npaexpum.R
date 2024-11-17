#' Title Extracting NPA Dataframe
#'
#' @param directory
#'
#' @return combined data
#' @export
#'
#'
npaexpump <- function(directory) {

  # List all Excel files in the specified directory
  excel_files <- list.files(path = directory, pattern = "\\.xlsx$", full.names = TRUE)

  # Define expected columns for consistency
  expected_columns <- c("NO.", "COMPANY", "REGULAR PETROL - RON 91 (GHS/Lt)",
                        "PREMIUM PETROL - RON 95 (GHS/Lt)", "DIESEL (GHS/Lt)",
                        "LPG (GHS/Kg)", "KEROSENE (GHS/Lt)", "MGO Local (GHS/Lt)",
                        "UNIFIED (GHp/Lt)")

  # Extract date from the file name
  extract_date <- function(file_name) {
    # Extract the date using a regular expression
    date_str <- sub(".*-(\\d{1,2}[a-zA-Z]{2}-[A-Za-z]+-\\d{4}).*", "\\1", basename(file_name), ignore.case = TRUE)

    # Remove ordinal suffixes and clean up the date string
    cleaned_date_str <- gsub("([0-9]+)(st|nd|rd|th)", "\\1", date_str, ignore.case = TRUE)
    cat("Extracted date string:", cleaned_date_str, "\n")

    # Attempt to parse the date using multiple formats
    parsed_date <- tryCatch(as.Date(cleaned_date_str, format = "%d-%B-%Y"), error = function(e) NA)
    if (is.na(parsed_date)) {
      parsed_date <- tryCatch(as.Date(cleaned_date_str, format = "%d-%b-%Y"), error = function(e) NA)
    }

    return(parsed_date)
  }

  # Read and clean individual Excel file
  read_excel_file <- function(file, is_first_file = FALSE) {
    # Extract the date from the file name
    date_value <- extract_date(file)

    # Read the Excel file using readxl
    if (is_first_file) {
      data <- readxl::read_excel(file, skip = 7, range = "B8:J1000")
    } else {
      data <- readxl::read_excel(file, skip = 8, range = "B9:J1000", col_names = FALSE)
      colnames(data) <- expected_columns
    }

    # Add the extracted date as a new column
    data$Date <- date_value

    # Ensure 'NO.' column is character type
    if ("NO." %in% colnames(data)) {
      data$`NO.` <- as.character(data$`NO.`)
    }

    # Convert all numeric columns, excluding 'NO.', 'COMPANY', and 'Date'
    data <- dplyr::mutate(data, dplyr::across(.cols = -c(`NO.`, COMPANY, Date),
                                              .fns = ~ as.numeric(gsub("[^0-9.]", "", .))))

    # Filter rows where 'NO.' is numeric
    data <- dplyr::filter(data, !is.na(`NO.`) & grepl("^[0-9]+$", as.character(`NO.`)))

    return(data)
  }

  # Process each file and combine the results
  combined_data <- NULL
  for (i in seq_along(excel_files)) {
    cat("\nProcessing file:", excel_files[i], "\n")

    # Read the file using the function
    new_data <- if (i == 1) {
      read_excel_file(excel_files[i], is_first_file = TRUE)
    } else {
      read_excel_file(excel_files[i], is_first_file = FALSE)
    }

    # Check if the Date column exists
    if (!"Date" %in% colnames(new_data) || all(is.na(new_data$Date))) {
      cat("Warning: Missing Date column for file:", excel_files[i], "\n")
    }

    # Ensure columns match the expected structure
    missing_cols <- setdiff(expected_columns, colnames(new_data))
    if (length(missing_cols) > 0) {
      new_data[missing_cols] <- NA
    }

    # Reorder columns to include the Date column
    new_data <- new_data[, c(expected_columns, "Date")]

    # Combine the data frames using dplyr::bind_rows
    combined_data <- if (is.null(combined_data)) {
      new_data
    } else {
      dplyr::bind_rows(combined_data, new_data)
    }

  }

  # Final adjustments
  # Move the Date column to be after the COMPANY column
  combined_data <- dplyr::relocate(combined_data, Date, .after = COMPANY)

  # Return the combined data frame
  return(combined_data)
}

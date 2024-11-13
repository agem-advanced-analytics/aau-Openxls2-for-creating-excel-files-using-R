get_data_position <-  function(title_data,
           row_info,
           main_data) {
    #' get data position
    #'
    #' get positions for title data, row info and main data
    #'
    #'
    #' @param title_data
    #' @param row_info
    #' @param main_data
    #'
    #' @return list containing 3 elements:
    #'  title_dimensions, row info dimennsions and main_dimensions.
 
  
    number_rows <- nrow(main_data)
    num_row_info_cols <- ncol(row_info)
    num_row_info_rows <- nrow(row_info)
    
    
    if (num_row_info_rows != number_rows) {
      stop(
        "row_info and main_data sections have different numbers of
        rows! Please fix this. "
      )
    }
    
    title_dimensions <-
      list(
        start_col = 1,
        start_row = 1,
        end_row = length(title_data),
        end_col = num_row_info_cols
      )
    
    
    start_row_data <- title_dimensions$end_row + 1 
    last_row_data <- start_row_data + num_row_info_rows
    
    row_info_dimensions <- list(
      start_col = 1,
      filter_col = 1,
      end_col = num_row_info_cols,
      start_row = start_row_data,
      end_row = last_row_data
    )
    
    start_col_main <- row_info_dimensions$end_col + 1
    end_col_main <- start_col_main + ncol(main_data) - 1
    main_dimensions <- list(
      start_col = start_col_main,
      end_col = end_col_main,
      start_row = start_row_data,
      end_row = last_row_data
    )
    
    return(
      list(
        title_dimensions = title_dimensions,
        row_info_dimensions = row_info_dimensions,
        main_dimensions = main_dimensions
      )
    )
  }

get_title_text <- function(sheet_info, date_to_print) {
  #' get title text
  #'
  #'create the data required for the title section
  #'
  #'
  #' @param sheet_info list containing information about sheet
  #' @param date_to_print string - most recent week included
  #'
  #' @return returns the desired content of the title section as a vector
  
  

  title_vector <-
    c(
      "Report demonstrating openxlsx2 functionalities",
      sheet_info$tab_title,
      date_to_print,
      sheet_info$tab_caveat,
      "Please see 'Missings and DQ issues' tab for more information."
    )
  
  return(title_vector)
  
}


create_rule <- function(section_data,
           conditional_formatting_type,
           conditional_formatting_detail) {
    #' create rule
    #'
    #' create rule for conditional formatting
    #'
    #' @param section_data data which will be conditionally formatted (tibble)
    #' @param conditional_formatting_type type of conditional formation
    #' either colorScale, arrows or traffic-
    #' others could be added but would need to add additional if statements
    #' @param conditional_formatting_detail extra information about conditional formatting
    #' for colourScale:
    #' - this is a vector of 2 or 3 colours as strings-
    #' -these will be the colours used in conditional formatting
    #' for arrows:
    #' - this will not be used- can be NA
    #' for traffic:
    #' - this should be "number" or "percent" based on whether the numbers to be
    #' conditionally formatted are numbers or percentages
    #'
    #' @return
    #' 

    
    if (conditional_formatting_type == "colorScale") {
      if (length(conditional_formatting_detail) == 3) {
        #if colorScale with 3 colours- rule is min, median and max of data

        section_vector <-
          section_data  |> t() |>  as.vector()## get into a vector
  

        min_point <- min(section_vector, na.rm = TRUE)    
        max_point <- max(section_vector, na.rm = TRUE)

        median_point <- median(section_vector, na.rm = TRUE)

        rule <- c(min_point, median_point, max_point)
        
      } else if (length(conditional_formatting_detail) == 2) {
        #if colorScale with 2 colours- rule is min and max of data
        
        section_vector <-
          section_data |> t() |>  as.vector()## get into a vector
        min_point <- min(section_vector, na.rm = TRUE)
        max_point <- max(section_vector, na.rm = TRUE)
        rule <- c(min_point, max_point)
      }
    } else if (conditional_formatting_type == "arrows") {
      #if arrows- rule is set to below
      rule <- c(-100, -50, 0 , 50, 100)
    } else if (conditional_formatting_type == "traffic") {
      rule <- c(0, 2, 4)
 
    }
    
    
    return(rule)
  }




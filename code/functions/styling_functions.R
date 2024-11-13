format_row_info <-  function(wb,
                             sheet_name,
                             start_col,
                             end_col,
                             row_info_rows ,
                             colour_to_use,
                             bold_var) {
  #' format row_info
  #'
  #' This function formats the row_info section of
  #' # the sheet
  #'
  #'
  #' @param wb workbook
  #' @param sheet_name stiring
  #' @param start_col int
  #' @param end_col int
  #' @param row_info_rows vector
  #' @param colour_to_use string
  #' @param bold_var "bold" if bold required
  #'
  #' @return workbook
  #'
  #'
  #get dimensions for row info section
  
  content_dims <- openxlsx2::wb_dims(
    cols = seq(from = start_col, to = end_col),
    rows = seq(
      from = min(row_info_rows),
      to = max(row_info_rows)
    )
  )

  
  #format row_info section
  wb <- wb |>
     openxlsx2::wb_add_font(
      sheet_name,
      dims = content_dims,
      bold = bold_var,
      name = "Calibri",
      size = 10
    ) |>
     openxlsx2::wb_add_fill(sheet_name,
                dims = content_dims,
                color = openxlsx2::wb_color(colour_to_use)) |>
     openxlsx2::wb_add_border(sheet_name, content_dims)
  
  return(wb)
}

add_blue_heading_formatting_with_total <-
  function(wb,
           sheet_name,
           start_col,
           end_col,
           start_row) {
    #' add blue heading formatting with total
    #'
    #' This makes the heading columns blue and makes the final column heading
    #' #of the section dark blue
    #'
    #' @param wb  workbook
    #' @param sheet_name sheet
    #' @param start_col col of the start of section (int)
    #' @param end_col col of the end of the section (int)
    #' @param start_row  row of the start of section( int)
    #'
    #' @return workbook
    main_headings_dim <-
       openxlsx2::wb_dims(cols = seq(from = start_col,
                         to = (end_col - 1)),
              rows = start_row)

    wb <- wb |>
       openxlsx2::wb_add_font(
        sheet_name,
        dims = main_headings_dim,
        bold = "bold",
        name = "Calibri",
        size = 10
      ) |>
      openxlsx2::wb_add_fill(sheet_name,
                  dims = main_headings_dim,
                  color = openxlsx2::wb_color("#DCE6F1"))  |>
       openxlsx2::wb_add_border(sheet_name,
                    dims = main_headings_dim) |>
       openxlsx2::wb_add_cell_style(dims = main_headings_dim, wrap_text = TRUE)

    total_headings_dim =  openxlsx2::wb_dims(cols = end_col,
                                 rows = start_row)
    wb <- wb |>
       openxlsx2::wb_add_font(
        sheet_name,
        dims = total_headings_dim,
        bold = "bold",
        color = openxlsx2::wb_color("white"),
        name = "Calibri",
        size = 10
      ) |>
       openxlsx2::wb_add_fill(sheet_name,
                  dims = total_headings_dim,
                  color = openxlsx2::wb_color("#16365C"))  |>
       openxlsx2::wb_add_border(sheet_name,
                    dims = total_headings_dim)

    return(wb)
  }

add_colour_headings <-
  function(wb,
           sheet_name,
           start_col,
           end_col,
           start_row,
           background_colour) {
    #' add color heading to all columns in main data section
    #'
    #' @param wb  workbook
    #' @param sheet_name sheet
    #' @param start_col col of the start of section (int)
    #' @param end_col col of the end of the section (int)
    #' @param start_row  row of the start of section( int)
    #' @param background_colour  colour of background 
    #'
    #' @return workbook
    main_headings_dim <-
      openxlsx2::wb_dims(cols = seq(from = start_col,
                                    to = end_col),
                         rows = start_row)
    
    wb <- wb |>
      openxlsx2::wb_add_font(
        sheet_name,
        dims = main_headings_dim,
        bold = "bold",
        name = "Calibri",
        size = 10
      ) |>
      openxlsx2::wb_add_fill(sheet_name,
                             dims = main_headings_dim,
                             color = openxlsx2::wb_color(background_colour))  |>
      openxlsx2::wb_add_border(sheet_name,
                               dims = main_headings_dim) |>
      openxlsx2::wb_add_cell_style(dims = main_headings_dim, wrap_text = TRUE)
    

    return(wb)
  }

add_comma_formatting <-
  function(wb, sheet_name, col_vector, row_vector) {
    #' Add comma formatting
    #'
    #'loops through col vector, adding comma formatting to specified rows
    #'
    #' @param wb workbook object
    #' @param sheet_name string
    #' @param col_vector vector of integers of columns to format
    #' @param row_vector vector of integers of rows to format (must be sequential)
    #'
    #' @return workbook
    for (col in col_vector) {
      dims_comma <-  openxlsx2::wb_dims(cols = col,
                              rows = row_vector)
      wb$add_numfmt(sheet_name, dims_comma, numfmt = "#,##0")
      
    }
    return(wb)
  }

add_conditional_formatting_section <- function(wb,
                                               sheet_name,
                                               start_col,
                                               end_col,
                                               rows,
                                               conditional_format_data,
                                               conditional_formatting_type,
                                               conditional_formatting_detail) {
  #' add conditional formatting section
  #'
  #' this function adds all the required conditional formatting to a section
  #' this involves applying the specified formatting to a whole section
  #' if is weekly is true, it will check if there is a qa issue,
  #' and apply qa specific formattng
  #'
  #' @param wb workbook
  #' @param sheet_name string
  #' @param start_col column to start conditional formatting
  #' @param end_col end column of conditional formatting
  #' @param rows vector of rows
  #' @param conditional_format_data dataframe of section
  #' @param conditional_formatting_type conditional formatting type (string)
  #' @param conditional_formatting_detail extra detail for con formatting (string)
  #'
  #' @return workbook

  conditional_rule <-
    create_rule(conditional_format_data,
                conditional_formatting_type,
                conditional_formatting_detail)

  wb <- wb |> add_specified_conditional_formatting(
    sheet_name,
    cols = seq(start_col, end_col),
    rows = rows,
    rule = conditional_rule,
    conditional_formatting_type =
      conditional_formatting_type,
    conditional_formatting_detail =
      conditional_formatting_detail
  )
    
  return(wb)
}


add_specified_conditional_formatting <-
  function(wb,
           sheet_name,
           cols,
           rows,
           rule,
           conditional_formatting_type,
           conditional_formatting_detail) {
    #' Add conditional formatting
    #'
    #' Adds conditional formatting to an area of a worksheet
    #'
    #' @param wb workbook
    #' @param cols vector of columns
    #' @param rows vector of rows
    #' @param rule rule for conditional formatted
    #'  (e.g. numbers for the different formats)
    #' @param conditional_formatting_type conditional formatting type (string)
    #' @param conditional_formatting_detail extra detail for con formatting (string)
    #' @param sheet_name sheet
    #'
    #' @return workbook


    area_dimensions = openxlsx2::wb_dims(cols = cols, rows = rows)

    if (conditional_formatting_type == "colorScale") {
      wb <- wb |> openxlsx2::wb_add_conditional_formatting(
        sheet_name,
        dims = area_dimensions,
        style = conditional_formatting_detail,
        rule = rule,
        #set this to be maximum
        type = conditional_formatting_type
      )
    } else if (conditional_formatting_type == "traffic") {
      wb <- wb |> openxlsx2::wb_add_conditional_formatting(
        sheet_name,
        dims = area_dimensions,
        rule = rule,
        #set this to be maximum
        type = "iconSet",
        params = list(iconSet = "3TrafficLights1",
                      reverse = TRUE)
      )

    } else if (conditional_formatting_type == "arrows") {
      wb <- wb |> openxlsx2::wb_add_conditional_formatting(
        sheet_name,
        dims = area_dimensions,
        rule = rule,
        #set this to be maximum
        type = "iconSet",
        params = list(iconSet = "5Arrows",
                      reverse = TRUE)
      )
    }
    return(wb)
  }


add_main_style <- function(wb,
                           sheet_name,
                           main_dimensions,
                           main_data, 
                           style_name) {
  
  #' add main style
  #'
  #' add style to a the main part of the workbook
  #' this function calls different functions depending on the style variable
  #'
  #' @param wb workbook
  #' @param sheet_name sheet
  #' @param main_dimensions dimensions of main section (list)
  #' @param main_data data for main section
  #' @param style_name style required (string)
  #'
  #' @return workbook
  #' 
  header_dimensions <-  openxlsx2::wb_dims(cols = seq(main_dimensions$start_col, main_dimensions$end_col),
                                rows = main_dimensions$start_row)
  numbers_dimensions <-   openxlsx2::wb_dims(cols = seq(main_dimensions$start_col, main_dimensions$end_col),
                                rows = seq(main_dimensions$start_row + 1,
                                            main_dimensions$end_row))  
  wb <- wb |>
     openxlsx2::wb_set_row_heights(sheet_name,
                       rows = main_dimensions$start_row,
                       heights = c(51)) |>
    openxlsx2::wb_add_font(
      sheet_name,
      dims = numbers_dimensions,
      name = "Calibri",
      size = 10
    ) |>
    openxlsx2::wb_add_border(sheet_name,
                             numbers_dimensions,
                             left_color = NULL,
                             right_color = NULL)  |>
    openxlsx2::wb_add_border(sheet_name,
                  header_dimensions,
                  inner_vgrid = "thin") |>
    openxlsx2::wb_add_cell_style(sheet_name,
                  dims = header_dimensions,
                  wrap_text = TRUE) |>
    openxlsx2::wb_add_cell_style(sheet_name,
                      dims = numbers_dimensions,
                      horizontal = "right") |>                  
    add_comma_formatting(sheet_name,
                         col_vector= seq(main_dimensions$start_col, main_dimensions$end_col),
                         row_vector = seq(main_dimensions$start_row, main_dimensions$end_row))
    
  
  if (style_name == "conditional_format_all") {
    wb <-
      wb |> add_main_style_conditional_format_all(
        sheet_name,
        main_dimensions,
        main_data
      )
    
  } else if (style_name == "months_with_total") {
    wb <-
      wb |> add_main_style_months_with_total(sheet_name,
                                         main_dimensions)
  } else{
    print(style_name)
    print("style function not valid")
    
  }
  
  return(wb)
}

add_main_style_conditional_format_all <-
  function(wb,
           sheet_name,
           main_dimensions,
           main_data) {
    #' Add main style departments
    #'
    #' This function adds the formatting for the a section of the main part of a
    #' #sheet where the style is "conditional_format_all"
    #'
    #' @param wb workbook
    #' @param sheet_name sheet
    #' @param main_dimensions dimensions of main section (list)
    #' @param main_data data for section - required for conditional formatting
    #' #, second ... section
    #' @return workbook
    wb <-
      wb |> add_colour_headings(sheet_name,
                                 start_col= main_dimensions$start_col,
                                 end_col =main_dimensions$end_col,
                                 start_row= main_dimensions$start_row,
                                 background_colour = "#D8E4BC")

    for (i in seq(1,7)){
      wb <- wb |>
      add_conditional_formatting_section(
        sheet_name = sheet_name,
        start_col =main_dimensions$start_col + i - 1 ,
        end_col = main_dimensions$start_col + i - 1,
        rows= seq(main_dimensions$start_row, main_dimensions$end_row),
        conditional_format_data = select(main_data, all_of(i)),
        conditional_formatting_type = "colorScale",
        conditional_formatting_detail = c("#FCFCFF", "blue")
      ) 
    }
    
    for (i in seq(8, 11)){
       wb<- wb |>
      add_conditional_formatting_section(
        sheet_name = sheet_name,
        start_col =main_dimensions$start_col + i - 1,
        end_col = main_dimensions$start_col + i - 1 ,
        rows= seq(main_dimensions$start_row, main_dimensions$end_row),
        conditional_format_data = select(main_data, all_of(i)),
        conditional_formatting_type = "traffic",
        conditional_formatting_detail = c(NA)
      )
    }
    
wb <- wb  |>
      wb_set_col_widths(
        sheet_name,
        cols = seq(from = main_dimensions$start_col, to =main_dimensions$end_col),
        widths = c(15)
      )
    return(wb)
  }

add_main_style_months_with_total <-
  function(wb,
           sheet_name,
           main_dimensions) {
    #' Add main style simple weeks
    #' #'
    #' This function adds the formatting for the a section of the main part of a
    #' #sheet where the style is "months_with_total"
    #'
    #' @param wb workbook
    #' @param sheet_name sheet
    #' @param main_dimensions dimensions of main section (list)
    #' @return workbook
    wb <- wb |>
      add_blue_heading_formatting_with_total(sheet_name,
                                             main_dimensions$start_col,
                                             main_dimensions$end_col,
                                             main_dimensions$start_row) |>
      openxlsx2::wb_set_col_widths(
        sheet_name,
        cols = seq(from = main_dimensions$start_col, to = main_dimensions$end_col - 1),
        widths = c(10)
      ) |>
      openxlsx2::wb_set_col_widths(sheet_name, cols = main_dimensions$end_col , widths = c(15))

    return(wb)
  }


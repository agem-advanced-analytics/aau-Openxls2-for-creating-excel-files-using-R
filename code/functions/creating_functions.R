

create_new_sheet <- function(wb, sheet_name, sheet_colour) {
  #' creates an empty sheet
  #'
  #' Calls function to make a sheet, taking in the workbook, data and
  #' # style requirements
  #'
  #' @param wb workbook
  #' @param sheet_name name of street (string)
  #' @param sheet_colour string
  #'
  #' @return workbook
  #

  
  wb <- openxlsx2::wb_add_worksheet(
    wb,
    sheet = sheet_name,
    tab_color = openxlsx2::wb_color(sheet_colour),
    grid_lines = FALSE
  )
  
  return(wb)
}

create_new_workbook <- function(template_path) {
  #' create workbook
  #'
  #' creates workbook, sets base font and adds required style
  #' can change options- data created , manager, etc
  #'
  #' @return workbook
  wb <- openxlsx2::wb_load(template_path) |>
    openxlsx2::wb_set_base_font(font_size = 10, font_name = "Calibri") 
  
  return(wb)
}

write_sheet <-  function(wb, sheet_info, data_list, date_to_print) {
    #' write sheet
    #'
    #' this function makes a sheet containing 3 parts- an title section,
    #' #row_info section and main body
    #'
    #' @param wb workbook
    #' @param sheet_info styling, sheet name and title for slide
    #' @param data_list list of the datasets for each parts
    #' @param date_to_print string representing date of data extract
    #'
    

    title_text <- get_title_text(sheet_info, date_to_print)
    
    style_info <- list(
      sheet_colour = sheet_info$tab_colour,
      sheet_style = sheet_info$sheet_style
    )
    
    
    dimensions <- get_data_position(
      title_data = title_text,
      row_info = data_list$row_info,
      main_data = data_list$main_data
    )
    

    #create and build up sheet
    wb <-
      create_new_sheet(wb, sheet_info$tab_name, style_info$sheet_colour) |>
      add_title_section(sheet_info$tab_name, dimensions$title_dimensions, title_text) |>
      add_row_info_section(sheet_info$tab_name,
                           dimensions$row_info_dimensions,
                           data_list$row_info) |>
      add_main_body(
        sheet_name = sheet_info$tab_name,
        main_data = data_list$main_data,
        main_dimensions = dimensions$main_dimensions,
        style_name = style_info$sheet_style
      )
    
    return(wb)
    
  }

write_section_content <- function(wb, sheet_name, data, dimensions) {
    #' Write section content
    #'
    #' This function writes a section of content to a spreadsheet
    #' @param wb workbook
    #' @param sheet_name name of sheet (character)
    #' @param data content to be added
    #' @param dimensions dimensions of section
    #'
    #' @return wb

    # add section text
    wb <- wb |>
      openxlsx2::wb_add_data(
        sheet_name,
        dims = openxlsx2::wb_dims(
          rows = seq(dimensions$start_row, dimensions$end_row),
          cols = seq(dimensions$start_col, dimensions$end_col)
        ),
        x = data,
        na.strings = "-"
      )
    
    return(wb)
    
  }

add_main_body <- function(wb, sheet_name, main_data, main_dimensions, style_name) {
  #'  add main body
  #'
  #' This function writes the main body of the sheet by looping through and
  #' # writing sections
  #'
  #' @param wb workbook
  #' @param sheet_name name of sheet (character)
  #' @param main_data main body data - nested dataframe with one row per section
  #' @param main_dimensions dimensions of main body
  #' @param style_name style required for section
  #'
  #' @return wb
  #'
  
  
  start_col <- main_dimensions$start_col

  
    #add section data, then add style for it
  wb <- wb |>
    write_section_content(sheet_name ,
                            main_data,
                            main_dimensions) |>
    add_main_style(
      sheet_name,
      main_dimensions,
      main_data,
      style_name
    )

  
  return(wb)
  
}

add_row_info_section <- function(wb, sheet_name, row_info_dimensions, row_info) {
    #' add row_info section
    #'
    #' add row_info and styling to a sheet based on row_info position list
    #' #(row_info_dimensions)
    #'
    #' @param wb workbook
    #' @param sheet_name sheet name (character)
    #' @param row_info_dimensions list containing row_info position information
    #' @param row_info row_infos
    #'
    #' @return none
    #' 
    #' 
    
    
    dims_top_row = wb_dims(
      cols = seq(
        row_info_dimensions$start_col,
        row_info_dimensions$end_col
      ),
      rows = row_info_dimensions$start_row
    )
    
    wb <- wb |>
      #add section data
      write_section_content(sheet_name ,
                            row_info,
                            row_info_dimensions) |>
      format_row_info(
        sheet_name,
        start_col = row_info_dimensions$start_col,
        end_col = row_info_dimensions$end_col,
        row_info_rows = (row_info_dimensions$start_row+1): row_info_dimensions$end_row,
        colour_to_use = "#F2F2F2",
        bold_var = ""
      ) |>
      #freeze row_info section
      openxlsx2::wb_freeze_pane(
        sheet_name,
        first_active_col = row_info_dimensions$end_col + 1,
        first_active_row = row_info_dimensions$start_row + 1
      ) |>
      # add filter to row_info section
      openxlsx2::wb_add_filter(
        sheet_name,
        rows = row_info_dimensions$start_row,
        cols = row_info_dimensions$filter_col
      ) |>
      # set column widths for row_info section
      openxlsx2::wb_set_col_widths(
        sheet_name,
        cols = seq(
          row_info_dimensions$start_col,
          row_info_dimensions$end_col
        ),
        widths = 20
      ) |>
      # add border for heading of row_info section
      openxlsx2::wb_add_border(
        sheet_name,
        dims = dims_top_row ,
        inner_vgrid = "thin"
      ) |>
      # add font and background for header or row info section
      openxlsx2::wb_add_fill(sheet_name,
                             dims = dims_top_row,
                             color = openxlsx2::wb_color("#BFBFBF")) |>
      openxlsx2::wb_add_border(sheet_name, dims_top_row) |>
      openxlsx2::wb_add_font(
        sheet_name,
        dims = dims_top_row,
        size = 10,
        name = "Calibri",
        bold= "bold",
        color= wb_color("black"))
    
    return(wb)
  }

add_title_section <-  function(wb, sheet_name,title_dimensions, title_data) {
  #' add title style
  #'
  #' adds the styling to a title section of a sheet in a workbook
  #'
  #' Style is different for first row, section row and remaining rows
  #'
  #' @param wb workbook
  #' @param sheet_name sheet_name (character)
  #' @param title_dimensions dimensions of title
  #'
  #' @return wb

  len_title <- title_dimensions$end_row - title_dimensions$start_row
  
  #add title section
  wb <- wb |>
    write_section_content(sheet_name ,
                          title_data,
                          title_dimensions) 
  

    # add formatting for top row
    dims_first <-
      openxlsx2::wb_dims(
        cols = seq(title_dimensions$start_col:title_dimensions$end_col),
        rows = title_dimensions$start_row
      )
    
    wb <- wb |>
      openxlsx2::wb_add_font(
        sheet = sheet_name,
        dims = dims_first,
        color = wb_color(name = "#FF0000"),
        size = 12,
        bold = "bold",
        name = "Calibri"
      ) 
  
    # add formatting to rows below top row
    row_sequence <-
      seq(from = (title_dimensions$start_row + 1),
          to = title_dimensions$end_row)
    
    dims_rest <-
      openxlsx2::wb_dims(
        cols = seq(title_dimensions$start_col:title_dimensions$end_col),
        rows = row_sequence
      )
    
    wb <- wb |>
      openxlsx2::wb_add_font(
        sheet = sheet_name,
        dims = dims_rest,
        size = 11,
        bold = "bold",
        name = "Calibri"
      ) 
  
  return(wb)
}

add_date_up_to<- function(wb, sheet_name, date_up_to, cell){
  #' add date up to value to a specific cell
  #'
  #' @param wb workbook
  #' @param sheet_name sheet_name (character)
  #' @param date_up_to date up to value (character)
  #' @param cell row_info to add data to (character e.g. "A4")
  #'
  #' @return wb
  wb<- wb |> openxlsx2::wb_add_data(sheet=sheet_name,
                          x= date_up_to,
                          dims= cell)
  return(wb)
}

add_tab_hyperlinks<- function(wb, sheet_name, report_sheets, column, start_row){
  #' add hyperlinks to every tab 
  #'
  #' @param wb workbook
  #' @param sheet_name sheet_name to add the hyperlinks to (character)
  #' @param report_sheets list containing all data and metadata for tabs
  #' @param column column to add the hyperlinks to
  #' @param start_row start row where hyperlinks should be added
  
  #'
  #' @return wb
  for (i in 1:length(report_sheets)){
    linked_sheet_name<- report_sheets[[i]]$sheet_params$tab_name
    link_dims<- openxlsx2::wb_dims(cols = column, 
                        rows = start_row + i - 1)
    
    desc_dims<- openxlsx2::wb_dims(cols= column + 2 ,
                        rows = start_row + i - 1)
    
    desc_text<- report_sheets[[i]]$sheet_params$tab_title
   
    link<- create_hyperlink(sheet = linked_sheet_name, text = linked_sheet_name) 
    wb <- wb |>
      openxlsx2::wb_add_formula(sheet_name, x = link, dims = link_dims) |>
      openxlsx2::wb_add_data(sheet_name, x= desc_text, dims= desc_dims)
  }
  return(wb)
  
  
}
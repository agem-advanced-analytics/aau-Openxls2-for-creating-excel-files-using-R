library(openxlsx2)
#dimensions_sepal_width<- wb_dims(rows=2:(nrow(my_data)+1), cols=2)
#dimensions_data <- wb_dims(rows=2:(nrow(my_data)+1), cols=1:ncol(my_data))



tab_name<- "first_tab"
my_data <- iris
dimensions_heading <- wb_dims(rows=1, cols=1:ncol(my_data))



## piping syntax- these functions do not modify in place
wb<-wb_workbook() 
wb<-wb |>
  wb_add_worksheet(tab_name,tab_color = wb_color("blue")) |>
  wb_add_data(tab_name,iris) |>
  wb_add_font(tab_name, dimensions_heading, name="Calibri", size = 11, bold = "single", color=wb_color("navy") ) |>
  wb_add_fill(tab_name, dimensions_heading,color = wb_color("beige")) |>
  wb_add_border(tab_name, dimensions_heading) |>
  wb_set_col_widths(tab_name, cols = 1:5, widths=c(10,10,10,10,20))
wb_save(wb,"outputs/simple_demo.xlsx", overwrite = TRUE)

## chaining syntax - these functions modify in place

wb<-wb_workbook() 

wb $ 
  add_worksheet(tab_name,tab_color = wb_color("blue")) $
  add_data(tab_name,iris)$
  add_font(tab_name, dimensions_heading, name="Calibri", size = 11, bold = "single", color=wb_color("navy") ) $
  add_fill(tab_name, dimensions_heading,color = wb_color("beige")) $
  add_border(tab_name, dimensions_heading) $
  set_col_widths(tab_name, cols = 1:5, widths=c(10,10,10,10,20))
wb$save("outputs/simple_demo.xlsx", overwrite = TRUE)







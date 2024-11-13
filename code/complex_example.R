library(openxlsx2)
library(dplyr)
library(readr)
library(readxl)
library(tidyr)
library(stringr)
library(lubridate)
library(forecast)

source("code/functions/add_main_style_functions.R")
source("code/functions/styling_functions.R")
source("code/functions/creating_functions.R")
source("code/functions/non_openxlsx2_functions.R")

mtcars_modified <-tibble::rownames_to_column(mtcars, "Car") |>
  mutate(Brand= str_split_i(Car, " ",1)) |>
  select(Brand, Car, mpg, cyl, disp, hp, drat, wt, qsec, vs, am, gear, carb)
air_passengers <- tibble( year = as.integer(time(AirPassengers)), 
        month = month.name[cycle(AirPassengers)],
        number_passengers= tsclean(AirPassengers)) |>
  pivot_wider(names_from = month, values_from = number_passengers) |>
  mutate(`Yearly Total` = rowSums(across(January:December)))


report_sheets<- list(sheet_1 = list(data = list(main_data = select(mtcars_modified, -Brand, -Car),
                                                row_info = select(mtcars_modified, Brand, Car)),
                                    sheet_params  = list(tab_name="Motor Trend Car Road Tests",
                                                         tab_title = "Road test results",
                                                         tab_caveat = "This is a built in R dataset",
                                                         sheet_color= "blue",
                                                         sheet_style="conditional_format_all")),
                     sheet_2 = list(data = list(main_data = select(air_passengers, -year),
                                                row_info = select(air_passengers, year)),
                                    sheet_params  = list(tab_name="Air Passengers over time",
                                                         tab_title = "Air Passenger results",
                                                         tab_caveat = "This is a built in R dataset",
                                                         sheet_color = "red",
                                                         sheet_style="months_with_total")))
                     

date_running <- format(today(),"%d %B %Y")


wb <- create_new_workbook("code/template_excel.xlsx") 
for (sheet in report_sheets) {
  wb <- wb |>
    write_sheet(
      sheet_info = sheet$sheet_params,
      data_list = sheet$data,
      date_to_print = date_running
    )
}

wb<- wb |>
  add_date_up_to("Front Sheet", date_running, "C5") |>
  add_tab_hyperlinks("Front Sheet", report_sheets,  column = 2, start_row = 14)

wb_save(wb, "outputs/complex_example.xlsx", overwrite = TRUE) 


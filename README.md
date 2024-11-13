# openxlsx2 Demonstration

This code was created for the purpose of demonstrating how openxlsx2 can be used to create a complex Excel report.

-   `code/`

    -   `basic_example.R`  - code demonstrating a simple use case of openxlsxs2

    -   `complex_example.R` - code demonstrating a complex use case of openxlsx2

    -   `convert_to_ods.R` - code to convert an .xlsx file to an .ods file.

    -   `template_excel.xlsx` - template excel file - with some formatting already added manually

    -   `functions/`

        -   `add_main_style_functions.R` - functions to style main body of excel

        -   `creating_functions.R` - functions to create sections, workbook and sheets

        -   `styling_functions.R` - helper functions to add styling to worksheets

        -   `non_openxlsxs2_functions.R` - helper functions that use data and parameters to get the dimensions of each part of the worksheet, create the text for the title section, and create the ranges for conditional formatting

-   `outputs/` - excel and ods outputs

-   `images/` - folder with screenshots and images used in presentation

-   `presentation.qmd` - quarto file to create presentation (simple version)

-   `template.pptx` - powerpoint template file (with AGEM AAU styling)

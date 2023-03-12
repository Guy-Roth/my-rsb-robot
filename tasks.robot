*** Settings ***
Documentation       Insert the sales data for the week and export it as a PDF.
...                 Added to GitHub repository.
...                 But had to manually update my name and email address?

Library             RPA.Browser.Selenium    auto_close=${FALSE}
Library             RPA.HTTP    ## to download files from remote web servers
Library             RPA.Excel.Files    ## to read the excel file w/o the app installed
Library             RPA.PDF


*** Variables ***
${DOWNLOAD_PATH}=       ${OUTPUT DIR}${/}downloads
## ${WORD_EXAMPLE}=    https://file-examples.com/wp-content/uploads/2017/02/file-sample_100kB.doc
## ${EXCEL_EXAMPLE}=    https://file-examples.com/wp-content/uploads/2017/02/file_example_XLS_10.xls


*** Tasks ***
Insert the sales data for the week and export it as a PDF
    Open the intranet website
    Log In
    Download the Excel file
    Fill the form using the data from the Excel file
    ## Fill and submit the form
    Collect the results
    Export the table as a PDF


*** Keywords ***
Open the intranet website
    Open Available Browser    https://robotsparebinindustries.com/

Log in
    Input Text    username    maria
    Input Password    password    thoushallnotpass
    Submit Form
    Wait Until Page Contains Element    id:sales-form

Fill and submit the form
    Input Text    firstname    John
    Input Text    lastname    Smith
    Input Text    salesresult    123
    Select From List By Value    salestarget    10000
    Click Button    Submit

Download the Excel file
    #    Download will go to the code directory by default.    Use target_file=/path to re-direct
    #    Download    https://robotsparebinindustries.com/SalesData.xlsx    overwrite=True    target_file=${DOWNLOAD_PATH}
    Download    https://robotsparebinindustries.com/SalesData.xlsx    overwrite=True

Fill the form using the data from the Excel file
    Open Workbook    SalesData.xlsx
    ${sales_reps}=    Read Worksheet As Table    header=True
    Close Workbook
    FOR    ${sales_rep}    IN    @{sales_reps}
        Fill and submit the form for one person    ${sales_rep}
    END

Fill and submit the form for one person
    [Arguments]    ${sales_rep}
    Input Text    firstname    ${sales_rep}[First Name]
    Input Text    lastname    ${sales_rep}[Last Name]
    Input Text    salesresult    ${sales_rep}[Sales]
    Select From List By Value    salestarget    ${sales_rep}[Sales Target]
    Click Button    Submit

Collect the results
    Screenshot    css:div.sales-summary    ${OUTPUT_DIR}${/}sales_summary.png
    #Screenshot    id.sales-results    ${OUTPUT_DIR}${/}sales_results.png
    TRY
        Screenshot    div:id.sales-results    ${OUTPUT_DIR}${/}sales_summary.png
        ##Fail    Element with locator
    EXCEPT    AS    ${error_message}
        Log To Console    ${error_message}
        Log    ${error_message}
        #Log To Console    Invalid div address. How to pass the actual exception text?
        #Log To Console    Error: Invalid div address
    ELSE
        Log To Console    ALL GOOD
    END

Export the table as a PDF
    Wait Until Element Is Visible    id:sales-results
    ${sales_results_html}=    Get Element Attribute    id:sales-results    outerHTML
    Html To Pdf    ${sales_results_html}    ${OUTPUT_DIR}${/}sales_results.pdf

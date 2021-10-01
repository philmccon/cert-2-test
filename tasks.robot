*** Settings ***
Documentation     Robot to enter weekly sales data into the RobotSpareBin Industries Intranet.
...               Saves the order HTML receipt as a PDF file.
...               Saves the screenshot of the ordered robot.
...               Embeds the screenshot of the robot to the PDF receipt.
...               Creates ZIP archive of the receipts and the images.
Library           RPA.Browser.Selenium
Library           RPA.Excel.Files
Library           RPA.Tables
Library           RPA.HTTP
Library           RPA.PDF
Library           RPA.Dialogs
Library           RPA.Archive
Library           RPA.Robocloud.Secrets
#  Library           RPA.Robocorp.Vault

*** Variables ***
${EXCEL_FILE_NAME}=    orders.csv
${EXCEL_FILE_URL}=     https://robotsparebinindustries.com/orders.csv
${Orders_URL}=         https://robotsparebinindustries.com/#/robot-order
${GLOBAL_RETRY_AMOUNT}=    10x
${GLOBAL_RETRY_INTERVAL}=    1s

*** Keywords ***
Get the URL from vault 
    ${url}=    Get Secret    cert2
    Log        ${url}

*** Keywords ***
Input form dialog
    Add heading       Alternative Excel File URL
    Add text input    url    label=URL address
    ...    label=web site address
    ...    placeholder=Enter URL here ${EXCEL_FILE_URL}
    ...    rows=1 
    ${result}=    Run dialog
    ${EXCEL_FILE_URL}=   Convert To String   ${result.url}  
    Log   ${EXCEL_FILE_URL}
#    ${secret}=          Get Secret      cert2
#    ${level}=           Set Log Level   NONE
#    Set To Dictionary   ${secret}       ${EXCEL_FILE_URL}
#    Set Log Level       ${level}
#    Set Secret          ${secret}

*** Keywords ***
Confirmation dialog
    Add icon      Warning
    Add heading   Do you want to use file ${EXCEL_FILE_URL}?
    Add submit buttons    buttons=No,Yes    default=Yes
    ${result}=    Run dialog
    IF   $result.submit == "No"
        Input form dialog
    END
    Log   ${EXCEL_FILE_URL}

*** Keywords ***
Get orders
    Download        ${EXCEL_FILE_URL}           overwrite=True
    ${table}=       Read Table From Csv       orders.csv      dialect=excel  header=True
    FOR     ${row}  IN  @{table}
        Log     ${row}
    END
    [Return]    ${table}

*** Keywords ***
Open the robot order website
    Open Available Browser    ${Orders_URL}

*** Keywords ***
Download an Excel file and read the rows
    Download        ${EXCEL_FILE_URL}           overwrite=True
    ${orders}=       Read table from CSV  ${EXCEL_FILE_NAME}   header=True
    [Return]         ${orders}

*** Keywords ***
Close the annoying modal
    Click Button    OK

*** Keywords ***
Preview the robot
    Click Element    id:preview
    Wait Until Element Is Visible    id:robot-preview

*** Keywords ***
Submit the order And Keep Checking Until Success
    Click Element    order
    Element Should Be Visible    xpath://div[@id="receipt"]/p[1]
    Element Should Be Visible    id:order-completion

*** Keywords ***
Go to order another robot
    Click Button    order-another

*** Keywords ***
Create a ZIP file of the receipts
    Archive Folder With Zip  ${CURDIR}${/}output${/}receipts   ${CURDIR}${/}output${/}receipt.zip

*** Keywords ***
Submit the order
    Wait Until Keyword Succeeds    ${GLOBAL_RETRY_AMOUNT}    ${GLOBAL_RETRY_INTERVAL}     Submit the order And Keep Checking Until Success

*** Keywords ***
Fill The Form 
   [Arguments]    ${localrow}
    ${head}=    Convert To Integer    ${localrow}[Head]
    ${body}=    Convert To Integer    ${localrow}[Body]
    ${legs}=    Convert To Integer    ${localrow}[Legs]
    ${address}=    Convert To String    ${localrow}[Address]
    Select From List By Value   id:head   ${head}
    Click Element   id-body-${body}
    Input Text      id:address    ${address}
    Input Text      xpath:/html/body/div/div/div[1]/div/div[1]/form/div[3]/input    ${legs}

*** Keywords ***
Collect The Results
    Screenshot    css:div.sales-summary    ${CURDIR}${/}output${/}sales_summary.png


*** Keywords ***
Store the receipt as a PDF file
    [Arguments]    ${order_number}
    Wait Until Element Is Visible    id:order-completion
    ${order_number}=    Get Text    xpath://div[@id="receipt"]/p[1]
    #Log    ${order_number}
    ${receipt_html}=    Get Element Attribute    id:order-completion    outerHTML
    Html To Pdf    ${receipt_html}    ${CURDIR}${/}output${/}receipts${/}${order_number}.pdf
    [Return]    ${CURDIR}${/}output${/}receipts${/}${order_number}.pdf

*** Keywords ***
Take a screenshot of the robot
    [Arguments]    ${order_number}
    Screenshot     id:robot-preview    ${CURDIR}${/}output${/}${order_number}.png
    [Return]       ${CURDIR}${/}output${/}${order_number}.png

*** Keywords ***
Embed the robot screenshot to the receipt PDF file
    [Arguments]    ${screenshot}   ${pdf}
    Open Pdf       ${pdf}
    Add Watermark Image To Pdf    ${screenshot}    ${pdf}
    Close Pdf      ${pdf}

*** Tasks ***
Order robots from RobotSpareBin Industries Inc
#    Get the URL from vault and Open the robot order website
#    Confirmation dialog
    Open the robot order website
    ${orders}=    Get orders
    FOR    ${row}    IN    @{orders}
        Close the annoying modal
        Fill the form    ${row}
        Preview the robot
        Submit the order
        ${pdf}=    Store the receipt as a PDF file    ${row}[Order number]
        ${screenshot}=    Take a screenshot of the robot    ${row}[Order number]
        Embed the robot screenshot to the receipt PDF file    ${screenshot}    ${pdf}
        Go to order another robot
    END
    Create a ZIP file of the receipts
    [Teardown]      Close Browser


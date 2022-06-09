*** Settings ***
Library    String
Library    openpyxl
Library    BuiltIn

*** Variables ***
${excel}    Excel_openpyxl.xlsx
${data}    write excel

*** Test Cases ***
Test Write Excel
    ${wb}    Load Workbook    ${CURDIR}/${excel}        #Load workbook คือการเปิดไฟล์ excel
    Log To Console    ${wb}
    ${ws}    Set Variable    ${wb['Sheet1']}            #set sheet1 ที่ต้องการจะเปิด
    Log To Console    ${ws}
    Evaluate    $ws.cell(2,3,"เพิ่มข้อความ")               #row 2 ,column 3 , เพิ่มข้อความลงไป
    Evaluate    $ws.cell(2,4,10)                        #row 2 ,column 4 , เพิ่มข้อความลงไป
    Evaluate    $ws.cell(2,5,'${data}')                 #row 2 ,column 5 , เพิ่มข้อความลงไป แบบใส่ตัวแปร
    Evaluate    $wb.save('${excel}')                    # save
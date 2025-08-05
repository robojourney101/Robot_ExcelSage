*** Settings ***
Library     ExcelSage
Library     BuiltIn

*** Variables ***
${EXCEL_PATH}     ${CURDIR}/data.xlsx
${SHEET_NAME}     Sheet1
${RESULT_COL}     C
${PERCENT_COL}    D

*** Test Cases ***
Read And Write Excel With ExcelSage
    [Documentation]    อ่านข้อมูลจากไฟล์ Excel, ประมวลผล และเขียนผลกลับลง Excel พร้อมสูตรเปอร์เซ็นต์

    # 🔹 เปิดไฟล์ Excel และเลือกแผ่นงาน
    Open Workbook          ${EXCEL_PATH}
    Set Active Sheet       ${SHEET_NAME}

    # 🔹 อ่านจำนวนแถวข้อมูลทั้งหมด
    ${row_count}=        Get Row Count    include_header=${True}

    # 🔄 วนลูปทีละแถว (เริ่มจาก row 2)
    ${end_row}=          Evaluate    ${row_count} + 1
    FOR    ${i}    IN RANGE    2     ${end_row} + 1
        ${name}=       Get Cell Value    A${i}
        ${score}=      Get Cell Value    B${i}
        ${score}=      Convert To Number    ${score}

        # ✅ เขียนผล "ผ่าน/ไม่ผ่าน" ลงคอลัมน์ C
        ${result}=     Run Keyword If    ${score} >= 50    Set Variable    ผ่าน
        ...            ELSE    Set Variable    ไม่ผ่าน
        Write To Cell      ${RESULT_COL}${i}    ${result}

        # ✅ เขียนสูตรเปอร์เซ็นต์ลงคอลัมน์ D เช่น =B2/100
        ${formula}=        Set Variable    =B${i}/100
        Write To Cell      ${PERCENT_COL}${i}    ${formula}

        Log     ${name} ได้คะแนน ${score} → ${result}    console=${True}
    END

    # 🔹 บันทึกไฟล์
    Save Workbook

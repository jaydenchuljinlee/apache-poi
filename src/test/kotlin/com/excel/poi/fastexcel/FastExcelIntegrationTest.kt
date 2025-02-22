package com.excel.poi.fastexcel

import com.excel.poi.common.dto.SheetInfo
import org.junit.jupiter.api.DisplayName
import org.junit.jupiter.api.Test
import java.io.FileOutputStream

class FastExcelIntegrationTest {
    @DisplayName("FastExcel 라이브러리를 통해 Excel을 출력한다.")
    @Test
    fun testFastExcel() {
        val outputStream = FileOutputStream("styled_output.xlsx")
        val workbook = FastExcelManager(outputStream)
        val sheet = workbook.createSheet("StyledSheet")

        val sheetInfo = SheetInfo.of(
            "sheet_1",
            listOf("ID", "이름", "이메일"),
            listOf(
                listOf("1", "홍길동", "hong@example.com"),
                listOf("2", "김철수", "kim@example.com"),
                listOf("3", "이영희", "lee@example.com")
            )
        )

        sheet.writeData(sheetInfo)

        val styleFactory = FastExcelStyleFactory(sheet.getSheet())
        styleFactory.applyHeaderStyle(0, 0, sheetInfo.headers.size - 1) // 헤더 스타일 적용
        styleFactory.applyDataStyle(1, sheetInfo.data.size, 0, sheetInfo.headers.size - 1) // 데이터 스타일 적용
        styleFactory.applyBorderStyle(0, sheetInfo.data.size, 0, sheetInfo.headers.size - 1) // 테두리 적용
        styleFactory.autoAdjustColumnWidth((0 until sheetInfo.headers.size).toList()) // 자동 열 너비 조정

        workbook.close()
    }
}
package com.excel.poi.business

import org.apache.poi.ss.usermodel.WorkbookFactory
import org.assertj.core.api.Assertions.assertThat
import org.junit.jupiter.api.BeforeEach
import org.junit.jupiter.api.DisplayName
import org.junit.jupiter.api.Test
import java.io.ByteArrayInputStream

class ExcelServiceIntegrationTest {
    private lateinit var excelService: ExcelService

    @BeforeEach
    fun setUp() {
        val excelExportService = ExcelExportService()
        excelService = ExcelService(excelExportService)
    }

    @DisplayName("엑셀 파일이 정상적으로 생성되는지 확인한다")
    @Test
    fun integrationTest() {
        // Given: 엑셀 파일 생성
        val excelData: ByteArray = excelService.generateExcel()

        // When: 생성된 파일을 Workbook으로 읽음
        ByteArrayInputStream(excelData).use { inputStream ->
            val workbook = WorkbookFactory.create(inputStream)

            // Then: 시트 개수 확인 (2개)
            assertThat(workbook.numberOfSheets).isEqualTo(2)

            val sheet1 = workbook.getSheetAt(0)
            val sheet2 = workbook.getSheetAt(1)

            // 시트 이름 확인
            assertThat(sheet1.sheetName).isEqualTo("Sheet1")
            assertThat(sheet2.sheetName).isEqualTo("Sheet2")

            // 헤더 확인 (10개 필드)
            val headerRow = sheet1.getRow(0)
            assertThat(headerRow.physicalNumberOfCells).isEqualTo(10)

            // 샘플 데이터 확인 (3개 로우)
            assertThat(sheet1.physicalNumberOfRows).isEqualTo(4) // 헤더 포함 4개

            val firstDataRow = sheet1.getRow(1)
            assertThat(firstDataRow.getCell(0).stringCellValue).isEqualTo("Data1-1")

            // 모든 열이 중앙 정렬인지 확인
            val cellStyle = firstDataRow.getCell(0).cellStyle
            assertThat(cellStyle.alignment).isEqualTo(org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER)
            assertThat(cellStyle.verticalAlignment).isEqualTo(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER)

            workbook.close()
        }
    }
}
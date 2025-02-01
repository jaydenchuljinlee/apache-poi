package com.excel.poi.business

import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.springframework.stereotype.Component
import java.io.ByteArrayOutputStream

@Component
class ExcelGenerator(
    private val excelStyleFactory: ExcelStyleFactory
) {
    fun createExcel(sheetNames: List<String>, headers: List<String>, data: List<List<String>>): ByteArray {
        val workbook: Workbook = XSSFWorkbook()

        // 각 시트 생성 및 데이터 추가
        sheetNames.forEach { sheetName ->
            val sheet = workbook.createSheet(sheetName)
            createSheetData(workbook, sheet, headers, data)
        }

        // ByteArray 변환
        ByteArrayOutputStream().use { outputStream ->
            workbook.write(outputStream)
            workbook.close()
            return outputStream.toByteArray()
        }
    }

    private fun createSheetData(workbook: Workbook, sheet: Sheet, headers: List<String>, data: List<List<String>>) {
        val headerStyle = excelStyleFactory.createHeaderStyle(workbook)
        val dataStyle = excelStyleFactory.createDataStyle(workbook)

        // 헤더 추가
        val headerRow = sheet.createRow(0)
        headers.forEachIndexed { colIndex, header ->
            val cell = headerRow.createCell(colIndex)
            cell.setCellValue(header)
            cell.cellStyle = headerStyle
        }

        // 데이터 추가
        data.forEachIndexed { rowIndex, rowData ->
            val row = sheet.createRow(rowIndex + 1)
            rowData.forEachIndexed { colIndex, value ->
                val cell = row.createCell(colIndex)
                cell.setCellValue(value)
                cell.cellStyle = dataStyle
            }
        }

        // 자동 열 너비 조정 및 행 높이 설정
        headers.indices.forEach { sheet.autoSizeColumn(it) }
        (0..data.size).forEach { sheet.getRow(it).heightInPoints = sheet.defaultRowHeightInPoints }
    }

}
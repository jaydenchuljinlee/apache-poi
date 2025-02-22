package com.excel.poi.poi.utils

import com.excel.poi.common.dto.SheetInfo
import org.apache.poi.ss.usermodel.*

data class WorkbookUtils(
    val workbook: Workbook,
    val sheet: Sheet
) {
    companion object {
        fun of(workbook: Workbook, sheetName: String): WorkbookUtils {
            return WorkbookUtils(
                workbook,
                workbook.createSheet(sheetName)
            )
        }

        fun of(workbook: Workbook, sheet: Sheet): WorkbookUtils {
            return WorkbookUtils(workbook, sheet)
        }
    }

    fun toSheetInfo(): SheetInfo {
        val dataList = mutableListOf<List<String>>()
        // 헤더 읽기
        val headerRow = sheet.getRow(0)
        val headers = (0 until headerRow.physicalNumberOfCells)
            .map { colIndex -> headerRow.getCell(colIndex).stringCellValue }

        // 데이터 읽기
        for (rowIndex in 1..sheet.lastRowNum) {
            val row = sheet.getRow(rowIndex)
            val rowData = mutableListOf<String>()
            headers.forEachIndexed { colIndex, header ->
                val cell = row.getCell(colIndex)
                val cellStr = cell?.toString() ?: "-"
                rowData.add(cellStr)
            }
            dataList.add(rowData)
        }
        return SheetInfo.of(sheet.sheetName, headers, dataList)
    }

    fun getHeaders(): List<String> {
        // 헤더 읽기
        val headerRow = sheet.getRow(0)
        return (0 until headerRow.physicalNumberOfCells)
            .map { colIndex -> headerRow.getCell(colIndex).stringCellValue }
    }


    fun create(sheetInfo: SheetInfo) {
        val headers = sheetInfo.headers
        val data = sheetInfo.data

        // 헤더 추가
        val headerStyle = createHeaderStyle(workbook)
        createHeader(sheetInfo.headers, headerStyle)

        // 데이터 추가
        val dataStyle = createDataStyle(workbook)
        createData(sheetInfo.data, dataStyle)

        // 자동 열 너비 조정 및 행 높이 설정
        headers.indices.forEach { sheet.autoSizeColumn(it) }
        (0..data.size).forEach { sheet.getRow(it).heightInPoints = sheet.defaultRowHeightInPoints }
    }

    private fun createHeader(headers: List<String>, style: CellStyle) {
        // 헤더 추가
        val headerRow = sheet.createRow(0)
        headers.forEachIndexed { colIndex, header ->
            val cell = headerRow.createCell(colIndex)
            cell.setCellValue(header)
            cell.cellStyle = style
        }
    }

    private fun createData(data: List<List<String>>, style: CellStyle) {
        data.forEachIndexed { rowIndex, rowData ->
            val row = sheet.createRow(rowIndex + 1)
            rowData.forEachIndexed { colIndex, value ->
                val cell = row.createCell(colIndex)
                cell.setCellValue(value)
                cell.cellStyle = style
            }
        }
    }

    private fun createHeaderStyle(workbook: Workbook): CellStyle {
        val style = workbook.createCellStyle()
        val font = workbook.createFont().apply {
            bold = true
            color = IndexedColors.WHITE.index
        }

        style.setFont(font)
        style.fillForegroundColor = IndexedColors.GREY_50_PERCENT.index
        style.fillPattern = FillPatternType.SOLID_FOREGROUND
        style.alignment = HorizontalAlignment.CENTER
        style.verticalAlignment = VerticalAlignment.CENTER
        return style
    }

    private fun createDataStyle(workbook: Workbook): CellStyle {
        val style = workbook.createCellStyle()
        style.alignment = HorizontalAlignment.CENTER
        style.verticalAlignment = VerticalAlignment.CENTER
        return style
    }
}
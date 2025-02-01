package com.excel.poi.business

import org.apache.poi.ss.usermodel.*
import org.springframework.stereotype.Component

@Component
class ExcelStyleFactory {
    fun createHeaderStyle(workbook: Workbook): CellStyle {
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

    fun createDataStyle(workbook: Workbook): CellStyle {
        val style = workbook.createCellStyle()
        style.alignment = HorizontalAlignment.CENTER
        style.verticalAlignment = VerticalAlignment.CENTER
        return style
    }

}
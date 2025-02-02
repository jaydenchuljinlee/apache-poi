package com.excel.poi.business

import com.excel.poi.dto.SheetInfo
import com.excel.poi.utils.WorkbookUtils
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.springframework.stereotype.Component
import java.io.ByteArrayOutputStream

@Component
class ExcelExportService {
    fun createExcel(sheets: List<SheetInfo>): ByteArray {
        val workbook: Workbook = XSSFWorkbook()

        // 각 시트 생성 및 데이터 추가
        sheets.forEach { sheetInfo ->
            val workbookUtils = WorkbookUtils.of(workbook, sheetInfo.name)
            workbookUtils.create(sheetInfo)
        }

        // ByteArray 변환
        ByteArrayOutputStream().use { outputStream ->
            workbook.write(outputStream)
            workbook.close()
            return outputStream.toByteArray()
        }
    }
}
package com.excel.poi.business

import com.excel.poi.dto.SheetInfo
import org.springframework.stereotype.Service

@Service
class ExcelService(
    private val excelExportService: ExcelExportService
) {
    fun generateExcel(): ByteArray {
        // 샘플 데이터 제공
        val headers = (1..10).map { "Field$it" }
        val data = listOf(
            (1..10).map { "Data1-$it" },
            (1..10).map { "Data2-$it" },
            (1..10).map { "Data3-$it" }
        )
        val sheets = mutableListOf<SheetInfo>()
        (1..2).forEach { idx ->
            sheets.add(SheetInfo.of("Sheet$idx", headers, data))
        }
        return excelExportService.createExcel(sheets)
    }
}
package com.excel.poi.business

import org.springframework.stereotype.Service

@Service
class ExcelService(
    private val excelGenerator: ExcelGenerator
) {
    fun generateExcel(): ByteArray {
        // 샘플 데이터 제공
        val headers = (1..10).map { "Field$it" }
        val data = listOf(
            (1..10).map { "Data1-$it" },
            (1..10).map { "Data2-$it" },
            (1..10).map { "Data3-$it" }
        )

        return excelGenerator.createExcel(listOf("Sheet1", "Sheet2"), headers, data)
    }
}
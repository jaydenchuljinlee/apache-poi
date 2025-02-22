package com.excel.poi.fastexcel

import com.excel.poi.common.dto.SheetInfo
import org.dhatim.fastexcel.Worksheet

class FastExcelWriter(private val sheet: Worksheet) {

    fun writeData(sheetInfo: SheetInfo) {
        // 헤더 작성
        sheetInfo.headers.forEachIndexed { colIndex, header ->
            sheet.value(0, colIndex, header)
        }

        // 데이터 작성
        sheetInfo.data.forEachIndexed { rowIndex, rowData ->
            rowData.forEachIndexed { colIndex, value ->
                sheet.value(rowIndex + 1, colIndex, value)
            }
        }
    }

    fun getSheet(): Worksheet {
        return sheet
    }
}
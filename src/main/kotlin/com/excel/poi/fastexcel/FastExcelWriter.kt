package com.excel.poi.fastexcel

import com.excel.poi.common.dto.SheetInfo
import com.excel.poi.fastexcel.dto.SheetPosition
import org.dhatim.fastexcel.Worksheet

class FastExcelWriter(private val sheet: Worksheet) {
    fun writeTableData(position: SheetPosition, sheetInfo: SheetInfo) {
        // 헤더 작성
        sheetInfo.headers.forEachIndexed { colIndex, header ->
            sheet.value(position.rowIndex, position.colIndex + colIndex, header)
        }

        // 데이터 작성
        sheetInfo.data.forEachIndexed { rowIndex, rowData ->
            rowData.forEachIndexed { colIndex, value ->
                sheet.value(position.rowIndex + rowIndex + 1, position.colIndex + colIndex, value)
            }
        }

        // 테이블 추가 후 rowIndex 증가
        position.rowIndex += sheetInfo.data.size + 1
    }

    fun writeRow(position: SheetPosition, rowData: List<String>) {
        rowData.forEachIndexed { colIndex, value ->
            sheet.value(position.rowIndex, position.colIndex + colIndex, value)
        }

        // 다음 행으로 이동
        position.rowIndex++
    }

    fun writeEmptyRow(position: SheetPosition, colCount: Int) {
        for (colIndex in 0 until colCount) {
            sheet.value(position.rowIndex, position.colIndex + colIndex, "") // 빈 값 입력
        }

        // 다음 행으로 이동
        position.rowIndex++
    }

    fun getSheet(): Worksheet {
        return sheet
    }
}
package com.excel.poi.fastexcel

import org.dhatim.fastexcel.BorderStyle
import org.dhatim.fastexcel.Worksheet

class FastExcelStyleFactory(private val sheet: Worksheet) {

    /** ✅ 헤더 스타일 적용 (볼드체 + 회색 배경) */
    fun applyHeaderStyle(startRow: Int, startCol: Int, endCol: Int) {
        for (col in startCol..endCol) {
            sheet.style(startRow, col)
                .fontName("Arial")
                .fontSize(12)
                .bold()
                .fillColor("CCCCCC") // 회색 배경색
                .horizontalAlignment("center")
                .set() // 스타일 적용
        }
    }

    /** ✅ 데이터 스타일 적용 (중앙 정렬) */
    fun applyDataStyle(startRow: Int, endRow: Int, startCol: Int, endCol: Int) {
        for (row in startRow..endRow) {
            for (col in startCol..endCol) {
                sheet.style(row, col)
                    .fontName("Arial")
                    .fontSize(11)
                    .horizontalAlignment("center")
                    .set() // 스타일 적용
            }
        }
    }

    /** ✅ 행 높이 및 열 너비 자동 조정 */
    fun autoAdjustColumnWidth(columnIndexes: List<Int>) {
        columnIndexes.forEach { sheet.width(it, 20.0) } // 기본 20
    }

    /** ✅ 특정 범위 병합 */
    fun mergeCells(startRow: Int, startCol: Int, endRow: Int, endCol: Int) {
        // 첫 번째 셀만 값을 유지하고 나머지는 공백으로 설정
        val mergedValue = sheet.value(startRow, startCol) ?: ""
        for (row in startRow..endRow) {
            for (col in startCol..endCol) {
                if (row == startRow && col == startCol) {
                    // 첫 번째 셀만 원래 값 유지
                    when (mergedValue) {
                        is String -> sheet.value(row, col, mergedValue)
                        is Number -> sheet.value(row, col, mergedValue)
                        is Boolean -> sheet.value(row, col, mergedValue)
                        is java.util.Date -> sheet.value(row, col, mergedValue)
                        is java.time.LocalDateTime -> sheet.value(row, col, mergedValue)
                        is java.time.LocalDate -> sheet.value(row, col, mergedValue)
                        is java.time.ZonedDateTime -> sheet.value(row, col, mergedValue)
                        else -> sheet.value(row, col, "") // 지원되지 않는 타입은 공백 처리
                    }
                } else {
                    sheet.value(row, col, "") // 나머지 셀은 공백
                }
            }
        }

        // 스타일 적용 (중앙 정렬 & 테두리 제거)
        for (row in startRow..endRow) {
            for (col in startCol..endCol) {
                sheet.style(row, col)
                    .horizontalAlignment("center")
                    .verticalAlignment("center")
                    .set()
            }
        }
    }

    /** ✅ 테두리 적용 */
    fun applyBorderStyle(startRow: Int, endRow: Int, startCol: Int, endCol: Int) {
        for (row in startRow..endRow) {
            for (col in startCol..endCol) {
                sheet.style(row, col)
                    .borderStyle(BorderStyle.THIN)
                    .borderColor("000000") // 검은색 테두리
                    .set()
            }
        }
    }
}
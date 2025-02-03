package com.excel.poi.business

import com.excel.poi.dto.SheetInfo
import com.excel.poi.utils.WorkbookUtils
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.springframework.stereotype.Component
import org.springframework.web.multipart.MultipartFile

@Component
class ExcelUploadService {
    fun readExcelFile(file: MultipartFile): List<SheetInfo> {
        val results = mutableListOf<SheetInfo>()

        // 엑셀 파일 파싱
        file.inputStream.use { inputStream ->
            val workbook = WorkbookFactory.create(inputStream)

            (0 .. workbook.numberOfSheets).map { idx ->
                val sheet = workbook.getSheetAt(idx) // N 번째 시트 읽음

                val workbookUtils = WorkbookUtils.of(workbook, sheet)
                results.add(workbookUtils.toSheetInfo())
            }

            workbook.close()
        }

        return results
    }

}
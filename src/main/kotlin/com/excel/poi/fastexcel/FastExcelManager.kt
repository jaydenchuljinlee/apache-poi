package com.excel.poi.fastexcel

import org.dhatim.fastexcel.Workbook
import java.io.OutputStream

class FastExcelManager(private val outputStream: OutputStream) {
    private val workbook = Workbook(outputStream, "MyApp", "1.0")
    private val sheets = mutableListOf<FastExcelWriter>()

    fun createSheet(sheetName: String): FastExcelWriter {
        val sheet = workbook.newWorksheet(sheetName)
        val writer = FastExcelWriter(sheet)
        sheets.add(writer)
        return writer
    }


    fun close() {
        workbook.finish()
    }
}
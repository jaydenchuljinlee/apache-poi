package com.excel.poi.dto

data class SheetInfo(
    val name: String,
    val headers: List<String>,
    val data: List<List<String>>
) {
    companion object {
        fun of(name: String, headers: List<String>, data: List<List<String>>): SheetInfo {
            return SheetInfo(name, headers, data)
        }
    }
}
package xyz.sinrin.excelbuilder

import org.apache.poi.ss.usermodel.Sheet

class ExcelReader<out T> internal constructor(private val clazz: Class<T>,
                         private val templateSheet: Sheet,
                         private val propertyNameIndex: Map<Int, String>,
                         private val firstDataRowIndex: Int,
                         private val lastDataRowIndex: Int
) {

    fun readSheet(): Iterable<T> {
        val wrapper = ExcelReaderWrapper(clazz).also { it.propertyNameIndex = this.propertyNameIndex }
        return if (lastDataRowIndex > 0) wrapper.readRows(templateSheet, firstDataRowIndex, lastDataRowIndex)
        else wrapper.readRows(templateSheet, firstDataRowIndex)
    }
}
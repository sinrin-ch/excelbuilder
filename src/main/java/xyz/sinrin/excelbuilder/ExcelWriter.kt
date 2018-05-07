package xyz.sinrin.excelbuilder

import org.apache.poi.ss.usermodel.Sheet

class ExcelWriter<T> internal constructor(private val clazz: Class<T>,
                     private val templateSheet: Sheet,
                     private val propertyNameIndex: Map<Int, String>,
                     private val firstDataRowIndex: Int,
                     private val styleToBeUsedRowIndex: Int,
                     private val dynamicColPredicate: ((T) -> Map<Int, Any?>)?) {

    fun writeSheet(excelDataObjects: Iterable<T>): Sheet {
        val wrapper = ExcelWriterWrapper(clazz)
                .also {
                    it.propertyNameIndex = this.propertyNameIndex
                    it.dynamicColPredicate = this.dynamicColPredicate
                }
        wrapper.writeRows(templateSheet, excelDataObjects, firstDataRowIndex, styleToBeUsedRowIndex)
        return this.templateSheet
    }

    fun writeSheet(excelDataArray: Array<T>): Sheet {
        return this.writeSheet(excelDataArray.asIterable())
    }
}
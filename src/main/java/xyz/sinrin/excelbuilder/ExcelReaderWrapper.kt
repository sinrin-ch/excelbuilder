package xyz.sinrin.excelbuilder

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import java.lang.reflect.Field
import java.lang.reflect.Method
import java.math.BigDecimal
import java.util.*

class ExcelReaderWrapper<T>(private val clazz: Class<T>) : ExcelBaseWrapper<T>(clazz) {
    /**
     * 从sheet读取数据并封装成list
     */
    fun readRows(sheet: Sheet, firstDataRowIndex: Int, lastDataRowIndex: Int): Iterable<T> {
        return (firstDataRowIndex..lastDataRowIndex)
                .map { this.convertToObject(sheet.getRow(it)) }
    }

    /**
     * 从sheet读取数据并封装成list
     */
    fun readRows(sheet: Sheet, firstDataRowIndex: Int): Iterable<T> {
        return this.readRows(sheet, firstDataRowIndex, sheet.lastRowNum)
    }

    /**
     * excel的行转换成object对象
     */
    private fun convertToObject(row: Row): T {
        val excelObject: T = clazz.newInstance()
        for (propertyNameIndex: Map.Entry<Int, String> in super.propertyNameIndex) {
            val index = propertyNameIndex.key
            val fieldName: String = propertyNameIndex.value
            val cell: Cell = row.getCell(index) ?: continue
            when {
                super.fieldNamesPropertyInfoDic.containsKey(fieldName) -> this.setCellValueToField(excelObject, fieldNamesPropertyInfoDic[fieldName]!!, cell)
                super.setterMethodDic.containsKey(fieldName) -> this.setCellValueBySetter(excelObject, setterMethodDic[fieldName]!!, cell)
            }
        }
        return excelObject
    }

    private fun setCellValueToField(excelObject: T, field: Field, cell: Cell) {
        val cellValue = this.readCellValue(cell, field.type) ?: return
        field.set(excelObject, cellValue)
    }

    private fun setCellValueBySetter(excelObject: T, method: Method, cell: Cell) {
        val cellValue = this.readCellValue(cell, method.parameterTypes.first()) ?: return
        method.invoke(excelObject, cellValue)
    }

    private fun readCellValue(cell: Cell, type: Class<*>): Any? {
        return when (type) {
            String::class.java ->
                when {
                    cell.cellType == 2 -> cell.numericCellValue.toString()
                    cell.cellType == 1 -> cell.booleanCellValue.toString()
                    cell.cellType == 3 -> ""
                    else -> cell.stringCellValue
                }
            Date::class.java -> cell.dateCellValue
            Float::class.java,
            Double::class.java,
            BigDecimal::class.java,
            Int::class.java,
            Byte::class.java,
            Short::class.java,
            Long::class.java -> cell.numericCellValue
            Boolean::class.java -> cell.booleanCellValue
            Char::class.java -> cell.stringCellValue.first()
            else -> null
        }
    }

}

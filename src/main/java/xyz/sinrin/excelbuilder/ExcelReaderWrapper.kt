package xyz.sinrin.excelbuilder

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import java.lang.reflect.Field
import java.lang.reflect.Method
import java.math.BigDecimal
import java.time.LocalDate
import java.time.ZoneId
import java.util.*

class ExcelReaderWrapper<T> internal constructor(private val clazz: Class<T>) : ExcelBaseWrapper<T>(clazz) {
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
                    cell.cellType == Cell.CELL_TYPE_NUMERIC -> cell.numericCellValue.toString()
                    cell.cellType == Cell.CELL_TYPE_BOOLEAN -> cell.booleanCellValue.toString()
                    cell.cellType == Cell.CELL_TYPE_BLANK -> ""
                    else -> cell.stringCellValue
                }
            Date::class.java -> cell.dateCellValue
            LocalDate::class.java -> cell.dateCellValue?.toInstant()?.atZone(ZoneId.systemDefault())?.toLocalDate()
            Float::class.java,
            java.lang.Float::class.java -> cell.numericCellValue.toFloat()
            Double::class.java,
            java.lang.Double::class.java -> cell.numericCellValue
            BigDecimal::class.java -> cell.numericCellValue.toBigDecimal()
            Int::class.java,
            Integer::class.java -> cell.numericCellValue.toInt()
            java.lang.Byte::class.java,
            Byte::class.java -> cell.numericCellValue.toByte()
            java.lang.Short::class.java,
            Short::class.java -> cell.numericCellValue.toShort()
            java.lang.Long::class.java,
            Long::class.java -> cell.numericCellValue.toLong()
            java.lang.Boolean::class.java,
            Boolean::class.java -> cell.booleanCellValue
            java.lang.Character::class.java,
            Char::class.java -> cell.stringCellValue.firstOrNull()
            else -> null
        }
    }

}

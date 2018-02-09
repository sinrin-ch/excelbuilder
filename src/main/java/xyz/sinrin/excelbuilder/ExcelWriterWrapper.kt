package xyz.sinrin.excelbuilder

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import java.math.BigDecimal
import java.util.*

class ExcelWriterWrapper<T>(clazz: Class<T>) : ExcelBaseWrapper<T>(clazz) {

//    private val getterDic: Map<String, Method> = clazz.methods.map { Pair<String, Method>(it.)}
//    /**
//     * 设置属性名,已经顺序.按照数组的名称排列,设置填充单元格的位置
//     */
//    private fun setPropertyNamesMapper(propertyNames: Array<String?>):
//            Map<Int, String?> {
//        this.propertyNameIndex = propertyNames
//                .indices
//                .map { i -> Pair<Int, String?>(i, propertyNames[i]) }
//                .filter { !it.second.isNullOrBlank() }
//                .toMap<Int, String?>()
//        return this.propertyNameIndex
//    }

    /**
     * 把obj对象数组写入工作表中
     */
    fun writeRows(sheet: Sheet,
                  objects: Iterable<T>, firstDataRowIndex: Int,
                  styleToBeUsedRowIndex: Int = firstDataRowIndex) {
        super.rowCellStyleDic = this.readRowStyles(sheet, styleToBeUsedRowIndex)
        objects.withIndex().forEach {
            val i = it.index
            val row: Row = sheet.getRow(firstDataRowIndex + i) ?:
                    sheet.createRow(firstDataRowIndex + i)
            this.writeCells(row, it.value)
        }
    }

    /**
     * 把obj对象的属性值,写入row的单元格中
     */
    private fun writeCells(row: Row, obj: T) {
        if (super.propertyNameIndex.isNotEmpty() && (super.fieldNamesPropertyInfoDic.isNotEmpty() || super.getterMethodDic.isNotEmpty())) {
            this.propertyNameIndex
                    .map<Int, String?, Pair<Int, Any?>> {
                        val propertyValue: Any? =
                                when {
                                    fieldNamesPropertyInfoDic.containsKey(it.value) -> this.fieldNamesPropertyInfoDic[it.value]!!.get(obj)
                                    getterMethodDic.containsKey(it.value) -> getterMethodDic[it.value]!!.invoke(obj)
                                    else -> null
                                }
                        return@map it.key to propertyValue
                    }
                    .forEach {
                        val propertyValue: Any = it.second ?: return@forEach
                        val cell: Cell = row.getCell(it.first) ?: row.createCell(it.first)
                        setCellValue(cell, propertyValue)
                        this.rowCellStyleDic[it.first]?.also { cell.cellStyle = it }
                    }
        }
        if (super.dynamicColPredicate != null) {
            val dynamicColDictionary: Map<Int, Any?> = super.dynamicColPredicate!!.invoke(obj)
            dynamicColDictionary.forEach {
                val col: Int = it.key
                val value: Any? = it.value
                val cell: Cell = row.getCell(col) ?: row.createCell(col)
                setCellValue(cell, value)
                this.rowCellStyleDic[col]?.also { cell.cellStyle = it }
            }
        }
    }

    /**
     * 从工作表中的第一个数据行读取单元格样式,根据单元格的列索引返回字典
     */
    private fun readRowStyles(sheet: Sheet, styleToBeUsedRowIndex: Int): Map<Int, CellStyle> {
        val row: Row = sheet.getRow(styleToBeUsedRowIndex) ?: return emptyMap()
        return row.map { Pair<Int, CellStyle>(it.columnIndex, it.cellStyle) }
                .toMap()
    }

    companion object {
        /**
         * 按不同的类型和值,写入单元格
         */
        fun setCellValue(newCell: Cell, value: Any?) {
            if (value == null) {
                return
            }
            when (value) {
                is String -> newCell.setCellValue(value)
                is Date -> newCell.setCellValue(value)
                is Double -> newCell.setCellValue(value)
                is Boolean -> newCell.setCellValue(value)
                is Char -> newCell.setCellValue(value.toString())
                is Float -> newCell.setCellValue(value.toDouble())
                is Int -> newCell.setCellValue(value.toDouble())
                is Byte -> newCell.setCellValue(value.toDouble())
                is Short -> newCell.setCellValue(value.toDouble())
                is Long -> newCell.setCellValue(value.toDouble())
                is BigDecimal -> newCell.setCellValue(value.toDouble())
            }
        }
    }
}
package xyz.sinrin.excelbuilder

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import java.lang.reflect.Field
import java.math.BigDecimal
import java.util.Date

class ExcelWriterWrapper<in T>(private val clazz: Class<T>) {
    // 属性名所对应的excel列的位置
    lateinit var propertyNameIndex: Map<Int, String?>
    //每一列的单元格样式.
    private lateinit var rowCellStyleDic: Map<Int, CellStyle>


    // 属性名和属性类型 字典
    private val fieldNamesPropertyInfoDic: Map<String, Field> = clazz.fields.map { Pair<String, Field>(it.name, it) }
            .toMap()

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
        this.rowCellStyleDic = this.readRowStyles(sheet, styleToBeUsedRowIndex)
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
    fun writeCells(row: Row, obj: T) {

        this.propertyNameIndex
                .filter { fieldNamesPropertyInfoDic.containsKey(it.value) }
                .forEach {
                    val field: Field = this.fieldNamesPropertyInfoDic[it.value]!!
                    val cell: Cell = row.getCell(it.key) ?: row.createCell(it.key)
                    val propertyValue = field.get(obj)
                    if (propertyValue != null) {
                        setCellValue(cell, propertyValue)
                        this.rowCellStyleDic[it.key]?.also { cell.cellStyle = it }
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
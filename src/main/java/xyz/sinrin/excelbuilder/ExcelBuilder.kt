package xyz.sinrin.excelbuilder

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File

class ExcelBuilder<T> {
    private val templateSheet: Sheet

    private var clazz: Class<T>

    // 设置列数,和属性名对应位置{{0, "Date"}, {1, "LineName"}}
    private var propertyNameIndex: Map<Int, String> = emptyMap()

    private var firstDataRowIndex: Int = 1

    private var styleToBeUsedRowIndex: Int = 1

    /**
     * 构建导出excel的creator
     */
    constructor(clazz: Class<T>, templateSheet: Sheet) {
        this.templateSheet = templateSheet
        this.clazz = clazz
    }

    constructor(clazz: Class<T>, templateWorkbook: Workbook) : this(clazz, templateWorkbook.getSheetAt(0))

    /**
     * 根据模板的文件名构建
     */
    constructor(clazz: Class<T>, templateWorkbookFileName: String, excelType: ExcelType) {
        val ins = File(templateWorkbookFileName).inputStream()
        val sheet = when (excelType) {
            ExcelBuilder.ExcelType.XLS -> HSSFWorkbook(ins)
            ExcelBuilder.ExcelType.XLSX -> XSSFWorkbook(ins)
        }.getSheetAt(0)
        this.templateSheet = sheet
        this.clazz = clazz
    }

    /**
     * 设置导入的excel data 的属性的顺序,通过字符串数组设置,如["id","name","","age"],表示第一列输出id,第二列输入name,空串则跳过
     */
    fun configPropertiesSort(vararg properties: String?): ExcelBuilder<T> {
        this.propertyNameIndex = properties
                .withIndex()
                .filter { it.value?.isNotBlank() == true }
                .map { Pair(it.index, it.value!!) }
                .toMap()
        return this
    }


    fun configFirstDataRowIndex(firstDataRowIndex: Int, styleToBeUsedRowIndex: Int = firstDataRowIndex): ExcelBuilder<T> {
        this.firstDataRowIndex = firstDataRowIndex
        this.styleToBeUsedRowIndex = styleToBeUsedRowIndex
        return this
    }

    fun buildWriter(): ExcelWriter<T> {
        return ExcelWriter(
                this.clazz,
                this.templateSheet,
                this.propertyNameIndex,
                this.firstDataRowIndex,
                this.styleToBeUsedRowIndex
        )
    }

    enum class ExcelType {
        XLS, XLSX
    }

    class ExcelWriter<in T>(private val clazz: Class<T>,
                            private val templateSheet: Sheet,
                            private val propertyNameIndex: Map<Int, String>,
                            private val firstDataRowIndex: Int,
                            private val styleToBeUsedRowIndex: Int) {

        fun writeSheet(excelDataObjects: Iterable<T>): Sheet {
            val wrapper = ExcelWriterWrapper(clazz)
                    .also { it.propertyNameIndex = this.propertyNameIndex }
            wrapper.writeRows(templateSheet, excelDataObjects, firstDataRowIndex, styleToBeUsedRowIndex)
            return this.templateSheet
        }
    }
}
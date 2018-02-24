package xyz.sinrin.excelbuilder

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.lang.reflect.Field

class ExcelBuilder<T> {
    private val templateSheet: Sheet

    private val clazz: Class<T>

    // 设置列数,和属性名对应位置{{0, "Date"}, {1, "LineName"}}
    private var propertyNameIndex: Map<Int, String> = emptyMap()

    private var firstDataRowIndex: Int = 1

    private var lastDataRowIndex: Int = -1

    private var styleToBeUsedRowIndex: Int = 1

    private var dynamicColPredicate: ((T) -> Map<Int, Any?>)? = null

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
            ExcelType.XLS -> HSSFWorkbook(ins)
            ExcelType.XLSX -> XSSFWorkbook(ins)
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

    fun configDynamicPropertiesCols(predicate: (T) -> Map<Int, Any?>): ExcelBuilder<T> {
        this.dynamicColPredicate = predicate
        return this
    }

    @JvmOverloads
    fun configDataRowIndex(firstDataRowIndex: Int, lastDataRowIndex: Int = -1): ExcelBuilder<T> {
        this.firstDataRowIndex = firstDataRowIndex
        this.lastDataRowIndex = lastDataRowIndex
        return this
    }

//    fun configDataRowIndex(firstDataRowIndex: Int): ExcelBuilder<T>{
//        this.firstDataRowIndex = firstDataRowIndex
//        return this
//    }

    fun configStyleToBeUsedRowIndex(styleToBeUsedRowIndex: Int): ExcelBuilder<T> {
        this.styleToBeUsedRowIndex = styleToBeUsedRowIndex
        return this
    }

    /**
     * 检测封装excel数据的实体类的注解
     */
    private fun checkAnnotation(): Map<Int, String> {
        if (this.propertyNameIndex.isNotEmpty()) {
            return this.propertyNameIndex
        }
//        val map = mutableMapOf<Int, String>()
//        clazz.declaredFields
//                .forEach { field ->
//                    val annotation = field.getAnnotation(CellConfig::class.java)
//                    if (annotation != null) {
//                        map.put(annotation.value, field.name)
//                    }
//                }
//        clazz.declaredFields
//                .forEach { field ->
//                    field.getAnnotation(CellConfig::class.java)?.let { map.put(it.value, field.name) }
//                }
        return clazz.declaredFields.mapNotNull<Field, Pair<Int, String>> { field ->
            val annotation: CellConfig = field.getAnnotation(CellConfig::class.java) ?: return@mapNotNull null
            return@mapNotNull annotation.value to field.name
        }.toMap<Int, String>()
    }

    fun buildWriter(): ExcelWriter<T> {
        // 如果没有指定属性名对应的列数,也没有指定动态委托,检测注解
        if (this.propertyNameIndex.isEmpty() && dynamicColPredicate == null) {
            this.propertyNameIndex = this.checkAnnotation()
        }
        return ExcelWriter(
                this.clazz,
                this.templateSheet,
                this.propertyNameIndex,
                this.firstDataRowIndex,
                this.styleToBeUsedRowIndex,
                this.dynamicColPredicate
        )
    }


}
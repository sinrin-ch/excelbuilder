package xyz.sinrin.excelbuilder

import org.apache.poi.ss.usermodel.CellStyle
import java.lang.reflect.Field
import java.lang.reflect.Method

abstract class ExcelBaseWrapper<T>(private val clazz: Class<T>) {
    // 属性名所对应的excel列的位置
    lateinit var propertyNameIndex: Map<Int, String>
    //每一列的单元格样式.
    protected lateinit var rowCellStyleDic: Map<Int, CellStyle>

    // 属性名和public属性类型 字典
    protected val fieldNamesPropertyInfoDic: Map<String, Field> by lazy {
        clazz.fields.map { Pair<String, Field>(it.name, it) }
                .toMap()
    }

    // getter方法字典
    protected val getterMethodDic: Map<String, Method> by lazy {
        clazz.methods
                .filter {
                    val name = it.name
                    return@filter name.startsWith("get") || name.startsWith("is")
                }
                .map {
                    val name = it.name
//                    val prefix = if (name.startsWith("get")) "get" else "is"
                    // 如果带有get,则去掉get.如果带有is,保留原样
                    val nameOri = if (name.startsWith("get")) name.removePrefix("get") else name
                    val firstLetter: Char = nameOri.first().toLowerCase()
                    val nameExtra: String = nameOri.substring(1)
                    return@map "$firstLetter$nameExtra" to it
                }
                .toMap()
    }

    // setter方法字典
    protected val setterMethodDic: Map<String, Method> by lazy {
        clazz.methods
                .filter { it.name.startsWith("set") }
                .map {
                    val methodName = it.name
                    val nameOri = methodName.removePrefix("set")
                    val firstLetter = nameOri.first().toLowerCase()
                    val fieldName = "$firstLetter${nameOri.substring(1)}"
                    val isFieldName = "is$nameOri"
                    return@map listOf<Pair<String, Method>>(fieldName to it, isFieldName to it)
                }.flatten().toMap()
    }

    //动态列的委托方法
    var dynamicColPredicate: ((T) -> Map<Int, Any?>)? = null
}
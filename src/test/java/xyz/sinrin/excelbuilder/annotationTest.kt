package xyz.sinrin.excelbuilder

import kotlin.test.assertNotEquals

fun main(args: Array<String>) {
    val clazz = User::class.java
    val map = mutableMapOf<Int,String>()
    clazz.declaredFields
            .forEach { field ->
                val annotation = field.getAnnotation(CellConfig::class.java)
                if (annotation != null) {
                    map.put(annotation.value,field.name)
                }
            }
    println(map)
}
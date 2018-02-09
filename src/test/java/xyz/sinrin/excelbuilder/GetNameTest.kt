package xyz.sinrin.excelbuilder

import java.lang.reflect.Method

fun main(args: Array<String>) {
    val clazz = User::class.java
    methodGetter(clazz).let { println(it) }
}

fun methodGetter(clazz: Class<*>): Map<String, Method> {
    return clazz.methods
            .filter {
                val name = it.name
                return@filter name.startsWith("get") || name.startsWith("is")
            }
            .map {
                val name = it.name
                val prefix = if (name.startsWith("get")) "get" else "is"
                val nameOri = name.removePrefix(prefix)
                val firstLetter: Char = nameOri.first().toLowerCase()
                val nameExtra: String = nameOri.substring(1)
                return@map "$firstLetter$nameExtra" to it
            }
            .toMap()
}
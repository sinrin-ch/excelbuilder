package xyz.sinrin.excelbuilder

fun main(args: Array<String>) {
    val m = listOf(1,2,null,3,4)
    m.mapNotNull { it?.times(2) }.forEach { println(it) }
}
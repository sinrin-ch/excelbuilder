package xyz.sinrin.excelbuilder

fun main(args: Array<String>) {
    val m = listOf(1, 2, null, 3, 4)
    m.mapNotNull<Int?, Pair<Int, Int>> {
        val a = it?:return@mapNotNull null
         a to a.times(2)
    }
            .forEach { println(it) }
}
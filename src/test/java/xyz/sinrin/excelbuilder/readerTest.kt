package xyz.sinrin.excelbuilder

fun main(args: Array<String>) {
    ExcelBuilder(Score::class.java, "/Users/sinrin/test.xlsx", ExcelType.XLSX)
            .configDataRowIndex(1, 5)
            .buildReader()
            .readSheet()
            .forEach { println(it) }

}
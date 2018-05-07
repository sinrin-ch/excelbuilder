package xyz.sinrin.excelbuilder

/**
 * Created by sinrin on 2018/5/5.
 */

fun main(args: Array<String>){
    val reader: ExcelReader<SpendTimeExcelModel> = ExcelBuilder(SpendTimeExcelModel::class.java, "/Users/sinrin/Downloads/数据库工序清单-整理.xlsx", ExcelType.XLSX)
            .configDataRowIndex(1)
            .buildReader()
    val sheet: Iterable<SpendTimeExcelModel> = reader.readSheet()
    sheet.forEach { println(it) }
}
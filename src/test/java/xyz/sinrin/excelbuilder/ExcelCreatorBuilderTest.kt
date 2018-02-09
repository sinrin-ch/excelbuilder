import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import xyz.sinrin.excelbuilder.ExcelBuilder
import xyz.sinrin.excelbuilder.User
import java.io.File

fun main(args: Array<String>) {
    val list = listOf<User>(User(1, "张三", 18),
            User(2, "李四", 19),
            User(3, "王五", 20))
    val user = User(3, "王五", 20)
    user.age
    val workbook: Workbook = XSSFWorkbook()
    val sheet: Sheet = workbook.createSheet()
    ExcelBuilder(User::class.java, sheet)
            .configPropertiesSort("id",null,"open", "name", "", "age")
            .configFirstDataRowIndex(3)
            .buildWriter()
            .writeSheet(list)
    val currentDir = System.getProperty("user.dir") +
            File.separator + "target"
    println(currentDir)
    File(currentDir).apply {
        if (!exists()) mkdirs()
    }
    val file: File = File(currentDir, "hehe.xlsx")
            .apply { if (!exists()) createNewFile() }
    file.outputStream().use(workbook::write)
}



package xyz.sinrin.excelbuilder;

import kotlin.jvm.functions.Function1;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class BuilderTest {
    public static void main(String[] args) throws IOException {

        List<User> list = new ArrayList<User>();
        list.add(new User(1, "张三", 18));
        list.add(new User(2, "李四", 19));
        list.add(new User(3, "王五", 20));

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        new ExcelBuilder<User>(User.class, sheet)
//                .configPropertiesSort("id", "name", "", "age")
                .configDynamicPropertiesCols(new Function1<User, Map<Integer, ?>>() {
                    @Override
                    public Map<Integer, ?> invoke(User user) {
                        Map<Integer, Object> map = new HashMap<Integer, Object>();
                        map.put(user.getId(), user.getId());
                        map.put(4, user.getName());
                        map.put(user.getAge() - 13, user.getAge());
                        return map;
                    }
                })
                .configDataRowIndex(3)
                .buildWriter()
                .writeSheet(
                        new User[]{
                                new User(1, "张三", 18),
                                new User(2, "李四", 19),
                                new User(3, "王五", 20)
                        });
        String curentDir = System.getProperty("user.dir") + File.separator + "target";
        File path = new File(curentDir);
        if (!path.exists()) {
            path.mkdirs();
        }

        File file = new File(curentDir, "hehe.xlsx");
        if (!file.exists()) {
            file.createNewFile();
        }

        OutputStream outputStream = new FileOutputStream(file);
        workbook.write(outputStream);
        outputStream.close();
    }
}

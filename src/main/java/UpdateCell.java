import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.Iterator;
import java.util.Locale;

public class UpdateCell {
    public static void main(String... args) throws IOException, InvalidFormatException {
        String result = "";
//        FileInputStream in = null;

        HSSFWorkbook wb = null;
        try {
//            in = new FileInputStream("Book1.xls");
            InputStreamReader in = new InputStreamReader(new FileInputStream("Book1.xls"), "CP1251");
            HSSFWorkbook myExcelBook = new HSSFWorkbook(new FileInputStream(file));
            Workbook workbook = WorkbookFactory.create(in);
            
            Sheet sheet = workbook.getSheetAt(0);
            Iterator<Row> it = sheet.iterator();
            while (it.hasNext()) {
                Row row = it.next();
                if (row.getRowNum() >= 7) {
                    Iterator<Cell> cells = row.iterator();
                    Cell cell = cells.next();
                    String cellValue = cell.getStringCellValue();
                    System.out.println(cellValue);
                    String s = "test";
                    int i = s.indexOf("e");
                    if (cellValue.contains("Фарба") || cellValue.contains("Краска")) {
                        System.out.println("Farba " + row.getRowNum());
                    } else if (cellValue.contains("Балон")) {
                        System.out.println("Balon " + row.getRowNum());
                    } else if (cellValue.trim().toLowerCase(Locale.ROOT).contains("Тексапон".trim().toLowerCase(Locale.ROOT))) {
                        System.out.println("Teksapon " + row.getRowNum());
                    }
                }
            }


            in.close();
//            FileOutputStream outputStream = new FileOutputStream("JavaBooksOutput.xls");
//            workbook.write(outputStream);
//            workbook.close();
//            outputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}


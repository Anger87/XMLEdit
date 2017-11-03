import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.Iterator;
import java.util.Locale;

public class UpdateCell {
    public static void main(String... args) throws IOException, InvalidFormatException {
        String result = "";
        FileInputStream in = null;

        try {
            in = new FileInputStream("/home/test/IdeaProjects/XMLEdit/XMLEdit/Book1.xls");
            Workbook workbook = WorkbookFactory.create(in);
            Sheet sheet = workbook.getSheetAt(0);
            Iterator<Row> it = sheet.iterator();
            while (it.hasNext()) {
                Row row = it.next();
                if (row.getRowNum() >= 7) {
                    Iterator<Cell> cells = row.iterator();
                    Cell cell = cells.next();
                    String cellValue = cell.getStringCellValue();
                    String importFlag = row.getCell(1).getStringCellValue();
                    System.out.println(cellValue);
//                    Фарба
                    if (cellValue.contains("Фарба") || cellValue.contains("Краска") || cellValue.contains("Лак") || cellValue.contains("грунт") || cellValue.contains("Морілка")) {
                        System.out.println("Farba " + row.getRowNum());
                        if (importFlag.contains("Импортированный товар")) {
                            row.getCell(10).setCellValue("+");
                        }
//                    Інші
                    } else if (cellValue.contains("Балон") || cellValue.contains("Диск") || cellValue.contains("Стрічка") || cellValue.contains("Пензлі") || cellValue.contains("Частини") || cellValue.contains("Пензлі")) {
                        System.out.println("Others " + row.getRowNum());
//                    Хімія
                    } else if (cellValue.contains("Тексапон") || cellValue.contains("Деріфат") || cellValue.contains("Дехікварт") || cellValue.contains("Трезаліт") || cellValue.contains("Розчинник") || cellValue.contains("Ларопал") || cellValue.contains("Глюкопон") || cellValue.contains("Трезоліт") || cellValue.contains("Шпаклівка") || cellValue.contains("ацетат") || cellValue.contains("Дехітон") || cellValue.contains("Тінувін") || cellValue.contains("Трилон") || cellValue.contains("Лютенсол")) {
                        System.out.println("Chemia " + row.getRowNum());
                    } else if (cellValue.length() > 1) {
                        result += cell.getStringCellValue() + " Row number: " + row.getRowNum() + "\n";

                    }
                }
            }

            System.out.println("not worked rows: " + "\n" + result);
            in.close();
            FileOutputStream outputStream = new FileOutputStream("JavaBooksOutput.xls");
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}


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
                int rowNum = row.getRowNum();
                //Set row number for formula
                int rowS = rowNum + 1;
                if (rowNum >= 7) {
                    Iterator<Cell> cells = row.iterator();
                    Cell cell = cells.next();
                    String name = cell.getStringCellValue();
                    String importFlag = row.getCell(1).getStringCellValue();
                    double sum = 0;
                    sum = row.getCell(4).getNumericCellValue();

                    if (name.length() > 0 && sum > 0) {
                        System.out.println(name + " / " + sum);
//                    Фарба
                        if (name.contains("Фарба") || name.contains("Краска") || name.contains("Лак") || name.contains("грунт") || name.contains("Морілка")) {
                            System.out.println("Farba " + rowNum);
                            if (importFlag.contains("Импортированный товар")) {
                                row.createCell(15).setCellValue("+");

                            }

//                    Інші
                        } else if (name.contains("Балон") || name.contains("Диск") || name.contains("Стрічка") || name.contains("Пензлі") || name.contains("Частини") || name.contains("Пензлі")) {
                            System.out.println("Others " + rowNum);
                            if (importFlag.contains("Импортированный товар")) {
                                row.createCell(19).setCellValue("+");
                                row.createCell(20).setCellFormula("J" +rowS);
                            } else {
                                row.createCell(27).setCellValue("+");
                                row.createCell(28).setCellFormula("J" +rowS);
                            }

//                    Хімія
                        } else if (name.contains("Тексапон") || name.contains("Деріфат") || name.contains("Дехікварт") || name.contains("Трезаліт") || name.contains("Розчинник") || name.contains("Ларопал") || name.contains("Глюкопон") || name.contains("Трезоліт") || name.contains("Шпаклівка") || name.contains("ацетат") || name.contains("Дехітон") || name.contains("Тінувін") || name.contains("Трилон") || name.contains("Лютенсол") || name.contains("Отверджувач") || name.contains("Антигравій")) {
                            System.out.println("Chemia " + rowNum);
                            if (importFlag.contains("Импортированный товар")) {
                                row.createCell(21).setCellValue("+");
                                row.createCell(22).setCellFormula("J" +rowS);
                            } else {
                                row.createCell(29).setCellValue("+");
                                row.createCell(30).setCellFormula("J" +rowS);
                            }

                        } else if (name.length() > 1) {
                            result += name + " Row number: " + row.getRowNum() + 1 + "\n";

                        }
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


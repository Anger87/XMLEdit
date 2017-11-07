package ExelLogic;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

public class UpdateCell {

    static String getPaintCount(String name) {
        if (!name.contains("Водорозчинний грунт лак")) {
            String result;
            String[] splited = name.split("\\s+");
            for (int i = 0; i < splited.length; i++) {
                if (splited[i].equals("L") || splited[i].equals("l") || splited[i].equals("л")) {
                    return splited[i - 1];
                } else if (splited[i].matches("\\d*" + "L") || splited[i].matches("\\d*" + "кг") || splited[i].matches("\\d*" + "л") || splited[i].matches("\\d*" + "," + "\\d*" + "л") || splited[i].matches("\\d*" + "," + "\\d*" + "кг")) {
                    result = splited[i].replaceAll("л|кг|L", "");
                    return result.replaceAll(",", ".");
                } else if (splited[i].matches("\\d*" + "мл")) {
                    result = splited[i].replaceAll("мл", "").replaceAll(",", ".");
                    double liters = Integer.parseInt(result);
                    result = String.valueOf(liters / 1000);
                    return result;
                }
            }
        }
        return "";
    }

    static double getSum(Row row) {
        try {

            int cellType = row.getCell(9).getCellType();
            if (cellType == Cell.CELL_TYPE_NUMERIC) {
                return row.getCell(9).getNumericCellValue();
            }
        } catch (NullPointerException e) {
            System.out.println("NullPointerException at RowNum: " + row.getRowNum());
        }
        return 0;
    }

    static String getNameCell(Cell cell) {
        try {

            int cellType = cell.getCellType();
            if (cellType == Cell.CELL_TYPE_STRING) {
                return cell.getStringCellValue();
            }
        } catch (NullPointerException e) {
            System.out.println("NullPointerException at Cell on Row: " + cell.getRowIndex());
        }
        return "";
    }

    public static void main(String... args) throws IOException, InvalidFormatException {
        String result = "";
        FileInputStream in = null;
        double sum;
        try {
//            in = new FileInputStream("/home/test/IdeaProjects/XMLEdit/XMLEdit/Book1.xls");
            in = new FileInputStream("/home/test/IdeaProjects/XMLEdit/XMLEdit/1-опт-остатки-3-кв.2017-281289.xls");
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

                    String name = getNameCell(cell);
                    if (!name.contains("Разом ")) {
                        sum = getSum(row);
                        if (name.length() > 1 && sum > 0) {
                            String importFlag = row.getCell(1).getStringCellValue();
//                            System.out.println(name + " / " + sum);
//                    Фарба
                            if (name.contains("фарба ") ||name.contains("Фарба ") || name.contains("Краска ") || name.contains("лак ")|| name.contains("Лак ") || name.contains("Грунт")|| name.contains("грунт") || name.contains("Морілка ")) {
//                                System.out.println(name + " | rowNum: " + rowNum + " | PaintCount: " + getPaintCount(name));
                                String paintCount = getPaintCount(name);
                                if (importFlag.contains("Импортированный товар")) {
                                    row.createCell(15).setCellValue("+");
                                    row.createCell(16).setCellFormula("J" + rowS);
                                    int rowQ = rowS + 1;
                                    if (paintCount.length() > 0)
                                        row.createCell(17).setCellFormula("J" + rowQ + "*" + paintCount);
                                    row.createCell(18).setCellFormula("R" + rowS + "/100");
                                } else {
                                    row.createCell(23).setCellValue("+");
                                    row.createCell(24).setCellFormula("J" + rowS);
                                    int rowQ = rowS + 1;
                                    if (paintCount.length() > 0)
                                        row.createCell(25).setCellFormula("J" + rowQ + "*" + paintCount);
                                    row.createCell(26).setCellFormula("Z" + rowS + "/100");
                                }

//                    Інші
                            } else if (name.contains("Балон") || name.contains("диск") || name.contains("Диск") || name.contains("Стрічка") || name.contains("Пензель ") || name.contains("Пензлі") || name.contains("Частини") || name.contains("Пензлі")) {
//                            System.out.println("Others " + rowNum);
                                if (importFlag.contains("Импортированный товар")) {
                                    row.createCell(19).setCellValue("+");
                                    row.createCell(20).setCellFormula("J" + rowS);
                                } else {
                                    row.createCell(27).setCellValue("+");
                                    row.createCell(28).setCellFormula("J" + rowS);
                                }

//                    Хімія
                            } else if (name.contains("Тексапон") || name.contains("Деріфат") || name.contains("Дехікварт") || name.contains("Трезаліт") || name.contains("Розчинник") || name.contains("Ларопал") || name.contains("Глюкопон") || name.contains("Трезоліт") || name.contains("Шпаклівка") || name.contains("ацетат") || name.contains("Дехітон") || name.contains("Тінувін") || name.contains("Трилон") || name.contains("Лютенсол") || name.contains("Отверджувач") || name.contains("Антигравій")) {
//                            System.out.println("Chemia " + rowNum);
                                if (importFlag.contains("Импортированный товар")) {
                                    row.createCell(21).setCellValue("+");
                                    row.createCell(22).setCellFormula("J" + rowS);
                                } else {
                                    row.createCell(29).setCellValue("+");
                                    row.createCell(30).setCellFormula("J" + rowS);
                                }

                            } else if (name.length() > 1) {
                                result += name + " Row number: " + row.getRowNum() + 1 + "\n";
                            }
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


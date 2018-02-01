package ExelLogic;

import Panel.Form;
import org.apache.commons.io.FileUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.Scanner;

public class UpdateCell {
    static Boolean importFlagProdaj;

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
                    // Coma case
                } else if (splited[i].matches("\\d*" + "L,") || splited[i].matches("\\d*" + "кг,") || splited[i].matches("\\d*" + "л,") || splited[i].matches("\\d*" + "," + "\\d*" + "л,") || splited[i].matches("\\d*" + "," + "\\d*" + "кг,")) {
                    result = splited[i].substring(0, splited[i].length() - 1);
                    result = result.replaceAll("л|кг|L", "");
                    return result.replaceAll(",", ".");
                } else if (splited[i].matches("\\d*" + "мл,")) {
                    result = splited[i].substring(0, splited[i].length() - 1);
                    result = result.replaceAll("мл", "").replaceAll(",", ".");
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

    static boolean checkIfNameContains(String filename, String name) throws IOException {
        String source = readFile(filename);
        String[] splitedName = name.split("\\s+");
        boolean isFound = false;
        for (String i : splitedName) {
            if (i.length() > 2 && source.indexOf(i) != -1) {
                isFound = true;
                break;
            }
        }
        return isFound;
    }


    private static String readFile(String pathname) throws IOException {

        File file = new File(pathname);
        StringBuilder fileContents = new StringBuilder((int) file.length());
        Scanner scanner = new Scanner(file);
        String lineSeparator = System.getProperty("line.separator");

        try {
            while (scanner.hasNextLine()) {
                fileContents.append(scanner.nextLine() + lineSeparator);
            }
            return fileContents.toString();
        } finally {
            scanner.close();
        }
    }

    public static void ScanDoc(String filePath) throws IOException, InvalidFormatException {
//    public static void main(String[] args) throws IOException, InvalidFormatException {
        String result = "";
        FileInputStream in = null;
        double sum;
        try {
            in = new FileInputStream(filePath);
//            Form.fileName = "Продажи 4 кв New.xls";
//            in = new FileInputStream(Form.fileName);
            Workbook workbook = WorkbookFactory.create(in);
            Sheet sheet = workbook.getSheetAt(0);
            Iterator<Row> it = sheet.iterator();
            while (it.hasNext()) {
                Row row = it.next();
                int rowNum = row.getRowNum();
                int rowS = rowNum + 1;
                Iterator<Cell> cells = row.iterator();
                // Logic for Prodaj
                if (Form.fileName.contains("родаж")) {
                    if (rowNum >= 12) {
                        Cell cell = cells.next();
                        String name = getNameCell(cell);
                        if (!name.contains("Підсумок")) {
                            setImportFlag(name);
                            sum = getSumProd(row);
                            if (name.length() > 1 && sum > 0) {
//                    Фарба
                                if (checkIfNameContains("paint.txt", name)) {
//                                System.out.println(name + " | rowNum: " + rowNum + " | PaintCount: " + getPaintCount(name));
                                    String paintCount = getPaintCount(name);
                                    // Import paint
                                    if (importFlagProdaj) {
                                        row.createCell(6).setCellValue("+");
                                        row.createCell(7).setCellFormula("F" + rowS);
                                        if (paintCount.length() > 0)
                                            row.createCell(8).setCellFormula("D" + rowS + "*" + paintCount);
                                        row.createCell(9).setCellFormula("I" + rowS + "/100");
                                    } else {
                                        row.createCell(14).setCellValue("+");
                                        row.createCell(15).setCellFormula("F" + rowS);
                                        if (paintCount.length() > 0)
                                            row.createCell(16).setCellFormula("D" + rowS + "*" + paintCount);
                                        row.createCell(17).setCellFormula("Q" + rowS + "/100");
                                    }

//                    Інші
                                } else if (checkIfNameContains("other.txt", name)) {
//                                System.out.println("Others " + rowNum);
                                    if (importFlagProdaj) {
                                        row.createCell(10).setCellValue("+");
                                        row.createCell(11).setCellFormula("F" + rowS);
                                    } else {
                                        row.createCell(18).setCellValue("+");
                                        row.createCell(19).setCellFormula("F" + rowS);
                                    }

//                    Хімія
                                } else if (checkIfNameContains("chemia.txt", name)) {
//                                System.out.println("Chemia " + rowNum);
                                    if (importFlagProdaj) {
                                        row.createCell(12).setCellValue("+");
                                        row.createCell(13).setCellFormula("F" + rowS);
                                    } else {
                                        row.createCell(20).setCellValue("+");
                                        row.createCell(21).setCellFormula("F" + rowS);
                                    }

                                } else if (name.length() > 1) {
                                    result += name + "\n";
                                }
                            }
                        }
                    }
                } else {
                    //Logic for Oborotky
                    if (rowNum >= 7) {
                        //Set row number for formula
                        Cell cell = cells.next();
                        String name = getNameCell(cell);
                        if (!name.contains("Разом ")) {
                            sum = getSum(row);
                            if (name.length() > 1 && sum > 0) {
                                String importFlag = row.getCell(1).getStringCellValue();
//                            System.out.println(name + " / " + sum);
//                    Фарба
                                if (checkIfNameContains("paint.txt", name)) {
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
                                } else if (checkIfNameContains("other.txt", name)) {
//                                System.out.println("Others " + rowNum);
                                    if (importFlag.contains("Импортированный товар")) {
                                        row.createCell(19).setCellValue("+");
                                        row.createCell(20).setCellFormula("J" + rowS);
                                    } else {
                                        row.createCell(27).setCellValue("+");
                                        row.createCell(28).setCellFormula("J" + rowS);
                                    }

//                    Хімія
                                } else if (checkIfNameContains("chemia.txt", name)) {
//                                System.out.println("Chemia " + rowNum);
                                    if (importFlag.contains("Импортированный товар")) {
                                        row.createCell(21).setCellValue("+");
                                        row.createCell(22).setCellFormula("J" + rowS);
                                    } else {
                                        row.createCell(29).setCellValue("+");
                                        row.createCell(30).setCellFormula("J" + rowS);
                                    }

                                } else if (name.length() > 1) {
                                    result += name + "\n";
                                }
                            }
                        }
                    }
                }
            }

            System.out.println("not worked rows: " + "\n" + result);
            in.close();
            FileOutputStream outputStream = new FileOutputStream(Form.fileName + "_Output.xls");
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();

            FileUtils.writeStringToFile(new File(Form.fileName + "_notScaned.txt"), result);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void setImportFlag(String name) {
        if (name.equals("Импортированный товар")) {
            importFlagProdaj = true;
        } else if (name.equals("Товар")) {
            importFlagProdaj = false;
        }
    }

    private static double getSumProd(Row row) {
        try {
            int cellType = row.getCell(5).getCellType();
            if (cellType == Cell.CELL_TYPE_FORMULA) {
                return row.getCell(5).getNumericCellValue();
            }
        } catch (NullPointerException e) {
            System.out.println("Sum not found at RowNum: " + row.getRowNum());
        }
        return 0;
    }

}


import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.write.*;
import jxl.write.Number;

import java.io.File;
import java.io.IOException;

import static jxl.Workbook.getWorkbook;

public class UpdateCellJXL {

    private static final String EXCEL_FILE_LOCATION = "/home/test/IdeaProjects/XMLEdit/XMLEdit/Book1.xls";

    public static void main(String[] args) {

        Workbook workbook = null;
        try {


            WorkbookSettings workbookSettings = new WorkbookSettings();
            workbookSettings.setEncoding("Cp1252");
//            workbook = getWorkbook(EXCEL_FILE_LOCATION, workbookSettings);

            Sheet sheet = workbook.getSheet(0);
            Cell cell1 = sheet.getCell(0, 0);
            System.out.print(cell1.getContents() + ":");    // Test Count + :
            Cell cell2 = sheet.getCell(0, 1);
            System.out.println(cell2.getContents());        // 1

            Cell cell3 = sheet.getCell(1, 0);
            System.out.print(cell3.getContents() + ":");    // Result + :
            Cell cell4 = sheet.getCell(1, 1);
            System.out.println(cell4.getContents());        // Passed

            System.out.print(cell1.getContents() + ":");    // Test Count + :
            cell2 = sheet.getCell(0, 2);
            System.out.println(cell2.getContents());        // 2

            System.out.print(cell3.getContents() + ":");    // Result + :
            cell4 = sheet.getCell(1, 2);
            System.out.println(cell4.getContents());        // Passed 2

        } catch (Exception e) {
            e.printStackTrace();
        } finally {

            if (workbook != null) {
                workbook.close();
            }

        }

    }

}


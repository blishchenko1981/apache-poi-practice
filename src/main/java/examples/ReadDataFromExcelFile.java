package examples;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

public class ReadDataFromExcelFile {

    public static void main(String[] args) throws IOException {
        String filePath = "data.xlsx";
        String sheetName = "data";

        InputStream in = new FileInputStream(filePath);
        Workbook workbook = WorkbookFactory.create(in);
        Sheet sheet = workbook.getSheet(sheetName);
        Row firstRow = sheet.getRow(0);
        Cell firstCell = firstRow.getCell(0);
        System.out.println("Value is: " + firstCell.getStringCellValue());

        Cell secondCell = firstRow.getCell(1);
        System.out.println("Value of 2nd cell is: " + secondCell.getStringCellValue());

        System.out.println("Number of cells in the row " + firstRow.getPhysicalNumberOfCells());

        firstRow.cellIterator().forEachRemaining(System.out::println);

        System.out.println("Number of rows in the sheet " + sheet.getPhysicalNumberOfRows());

        sheet.forEach(r -> {
            r.cellIterator().forEachRemaining(c -> System.out.print(c.toString()+ " "));
            System.out.println("");
        });
    }
}

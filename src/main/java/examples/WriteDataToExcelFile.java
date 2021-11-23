package examples;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class WriteDataToExcelFile {
    public static void main(String[] args) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("new sheet");

        Row firstRow = sheet.createRow(0);

        Cell firstCell = firstRow.createCell(0);
        firstCell.setCellValue("name");

        Cell secondCell = firstRow.createCell(1);
        secondCell.setCellValue("age");

        String path = "my_new_file.xlsx";

        try(FileOutputStream out = new FileOutputStream(path)){
            workbook.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

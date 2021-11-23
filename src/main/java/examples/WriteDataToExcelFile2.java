package examples;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class WriteDataToExcelFile2 {
    public static void main(String[] args) throws IOException {
        String filePath = "data.xlsx";
        FileInputStream inputStream = new FileInputStream(filePath);

        Workbook workbook = WorkbookFactory.create(inputStream);
        Sheet sheet = workbook.getSheetAt(0);

        sheet.getRow(0).createCell(6).setCellValue("status");
        sheet.getRow(1).createCell(6).setCellValue("passed");
        sheet.getRow(2).createCell(6).setCellValue("passed");
        sheet.getRow(3).createCell(6).setCellValue("failed");
        
        try(FileOutputStream out = new FileOutputStream(filePath)) {
            workbook.write(out);
        }catch (IOException e){
            e.printStackTrace();
        }
    }
}

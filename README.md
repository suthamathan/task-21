# task-21
task 21
Program to Write Data to Excel:




import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class WriteExcel {

    public static void main(String[] args) {
        try (Workbook workbook = new XSSFWorkbook()) {
            // Create a new sheet
            Sheet sheet = workbook.createSheet("Sheet1");

            // Write column headers
            Row headerRow = sheet.createRow(0);
            String[] headers = {"Name", "Age", "Email"};
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
            }

            // Write data to the sheet
            String[][] data = {
                    {"John Doe", "30", "john@test.com"},
                    {"Jane Doe", "28", "jane@test.com"},
                    {"Bob Smith", "35", "bob@test.com"},
                    {"Swapnil", "37", "swapnil@test.com"}
            };

            for (int i = 0; i < data.length; i++) {
                Row row = sheet.createRow(i + 1);
                for (int j = 0; j < data[i].length; j++) {
                    Cell cell = row.createCell(j);
                    cell.setCellValue(data[i][j]);
                }
            }

            // Save the workbook to a file
            try (FileOutputStream fileOut = new FileOutputStream("output.xlsx")) {
                workbook.write(fileOut);
                System.out.println("Data written to Excel file successfully.");
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}


Program to Read Data from Excel:

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class ReadExcel {

    public static void main(String[] args) {
        try (Workbook workbook = new XSSFWorkbook(new FileInputStream("output.xlsx"))) {
            // Get the first sheet
            Sheet sheet = workbook.getSheetAt(0);

            // Iterate through rows and columns and print the data
            for (Row row : sheet) {
                for (Cell cell : row) {
                    System.out.print(cell.getStringCellValue() + "\t");
                }
                System.out.println();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

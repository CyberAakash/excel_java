package excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

public class ExcelCreator {

    public static void main(String[] args) {
        // Create a new workbook
        try (Workbook workbook = new XSSFWorkbook()) {
            // Create a blank sheet
            Sheet sheet = workbook.createSheet("Sheet1");

            // Sample data (replace this with your actual data)
            Object[][] data = {
                    {"Column1", "Column2"},
                    {1, "A"},
                    {2, "B"},
                    {3, "C"},
                    {4, "D"},
                    {5, "E"}
            };

            // Populate the sheet with data
            for (int rowNum = 0; rowNum < data.length; rowNum++) {
                Row row = sheet.createRow(rowNum);
                for (int colNum = 0; colNum < data[rowNum].length; colNum++) {
                    Cell cell = row.createCell(colNum);
                    if (data[rowNum][colNum] instanceof String) {
                        cell.setCellValue((String) data[rowNum][colNum]);
                    } else if (data[rowNum][colNum] instanceof Integer) {
                        cell.setCellValue((Integer) data[rowNum][colNum]);
                    }
                }
            }

            // Generate filename with the specified format
            String tableName = "TableName";
            String currentDate = new java.text.SimpleDateFormat("ddMMyyyy_HHmmss").format(new Date());
            String filename = tableName + "_" + currentDate + ".xlsx";

            // Write the workbook to a file
            try (FileOutputStream fileOut = new FileOutputStream(filename)) {
                workbook.write(fileOut);
                System.out.println("Excel sheet \"" + filename + "\" created successfully.");
            } catch (IOException e) {
                e.printStackTrace();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

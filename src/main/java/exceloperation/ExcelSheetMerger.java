package exceloperation;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class ExcelSheetMerger {
    public static void main(String[] args) {
        // Paths to the input Excel files
        String[] inputFilePaths = {"example.xlsx", "example2.xlsx"};

        // Path to the output Excel file
        String outputFilePath = "output.xlsx";

        // Set to store unique data rows
        Set<List<String>> uniqueRows = new LinkedHashSet<>();

        try {
            // Read data from each input Excel file
            for (String inputFilePath : inputFilePaths) {
                FileInputStream fis = new FileInputStream(new File(inputFilePath));
                Workbook workbook = WorkbookFactory.create(fis);
                Sheet sheet = workbook.getSheetAt(0); // Assuming the data is in the first sheet

                // Iterate over each row of the sheet
                for (Row row : sheet) {
                    List<String> rowData = new ArrayList<>();
                    // Iterate over each cell of the row
                    for (Cell cell : row) {
                        // Get the cell value and add it to the rowData list
                        rowData.add(cell.toString());
                    }
                    // Add the rowData list to the set
                    uniqueRows.add(rowData);
                }

                // Close the workbook and FileInputStream
                workbook.close();
                fis.close();
            }

            // Write the unique data to a new Excel file
            writeDataToExcel(outputFilePath, uniqueRows);

            System.out.println("Excel sheets merged successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void writeDataToExcel(String outputFilePath, Set<List<String>> uniqueRows) throws IOException {
        // Create a new Workbook object for the output Excel file
        Workbook workbook = new XSSFWorkbook();
        // Create a new Sheet in the workbook
        Sheet sheet = workbook.createSheet("MergedSheet");

        int rowIndex = 0;
        // Iterate over each unique row in the set
        for (List<String> rowData : uniqueRows) {
            // Create a new Row in the sheet
            Row row = sheet.createRow(rowIndex++);
            int cellIndex = 0;
            // Iterate over each cell value in the rowData list
            for (String cellValue : rowData) {
                // Create a new Cell in the row and set its value
                Cell cell = row.createCell(cellIndex++);
                cell.setCellValue(cellValue);
            }
        }

        // Write the workbook to the output Excel file
        try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
            workbook.write(fos);
        }

        // Close the workbook
        workbook.close();
    }
}

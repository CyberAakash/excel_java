package exceloperation;

import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class ExcelReader {
    public static void main(String[] args) {
        // Path to the Excel file
        String filePath = "example.xlsx";

        try {
            // Create a FileInputStream to read the Excel file
            FileInputStream fis = new FileInputStream(new File(filePath));

            // Create a Workbook object for the Excel file
            Workbook workbook = WorkbookFactory.create(fis);

            // Get the first sheet of the workbook
            Sheet sheet = workbook.getSheetAt(0); // Assuming the data is in the first sheet

            // Iterate through each row of the sheet
            for (Row row : sheet) {
                // Check if it's the first row (header row)
                if (row.getRowNum() == 0) {
                    // Iterate through each cell of the header row and print the column headers
                    for (Cell cell : row) {
                        System.out.print(cell.toString() + "\t");
                    }
                    System.out.println(); // Move to the next line after printing the headers
                } else {
                    // Iterate through each cell of the data rows and print the cell values
                    for (Cell cell : row) {
                        System.out.print(cell.toString() + "\t");
                    }
                    System.out.println(); // Move to the next line after printing each data row
                }
            }

            // Close the workbook and FileInputStream
            workbook.close();
            fis.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

package exceloperation;

import org.apache.poi.ss.usermodel.*;
import org.jetbrains.annotations.NotNull;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.InputMismatchException;
import java.util.Scanner;

public class ExcelColumnDataPrinter {
    public static void main(String[] args) {
        // Path to the Excel file
        String filePath = "example.xlsx";

        Scanner scanner = new Scanner(System.in);

        try {
            // Create a FileInputStream to read the Excel file
            FileInputStream fis = new FileInputStream(new File(filePath));

            // Create a Workbook object for the Excel file
            Workbook workbook = WorkbookFactory.create(fis);

            // Get the first sheet of the workbook
            Sheet sheet = workbook.getSheetAt(0); // Assuming the data is in the first sheet

            // Prompt user for column index
            int columnIndex = getColumnIndexFromUser(scanner, sheet.getLastRowNum());

            // Validate column index
            if (columnIndex < 0 || columnIndex >= sheet.getRow(0).getLastCellNum()) {
                System.out.println("Invalid column index! Please enter a valid index within the range.");
                return;
            }

            // Iterate through each row of the sheet
            for (Row row : sheet) {
                // Check if the row is not null and if the column index is within the row's range
                if (row != null && columnIndex < row.getLastCellNum()) {
                    // Get the cell in the specified column
                    Cell cell = row.getCell(columnIndex);
                    // Print or process the data from the cell
                    if (cell != null) {
                        System.out.println(cell.toString());
                    } else {
                        System.out.println("Cell is empty");
                    }
                } else {
                    System.out.println("Column index is out of range for one or more rows.");
                }
            }

            // Close the workbook and FileInputStream
            workbook.close();
            fis.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // Method to prompt the user for the column index
    private static int getColumnIndexFromUser(@NotNull Scanner scanner, int maxIndex) {
        int columnIndex;
        while (true) {
            try {
                System.out.print("Enter the column index (0-based) to display data from: ");
                columnIndex = scanner.nextInt();
                scanner.nextLine(); // Consume newline character
                if (columnIndex < 0 || columnIndex >= maxIndex) {
                    System.out.println("Invalid column index! Please enter a valid index within the range.");
                } else {
                    break; // Exit the loop if the index is valid
                }
            } catch (InputMismatchException e) {
                System.out.println("Invalid input! Please enter a valid integer value.");
                scanner.nextLine(); // Consume the invalid input
            }
        }
        return columnIndex;
    }

}

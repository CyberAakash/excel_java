package exceloperation;

import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Scanner;

public class ExcelReaderByReg {
    public static void main(String[] args) {
        // Path to the Excel file
        String filePath = "example.xlsx";

        try (Scanner scanner = new Scanner(System.in);
             FileInputStream fis = new FileInputStream(new File(filePath));
             Workbook workbook = WorkbookFactory.create(fis)) {

            // Get the first sheet of the workbook
            Sheet sheet = workbook.getSheetAt(0); // Assuming the data is in the first sheet

            // Print message about the required format for registration number
            System.out.println("Please enter the registration number in the format 'SSx..x', where s represents character and x represents a digit.");

            // Prompt the user to enter the registration number to search
            String registrationNumber;
            boolean isValid = false;
            do {
                System.out.print("\nEnter Register Number to search: ");
                registrationNumber = scanner.next();

                // Validate the format and type of the registration number
                isValid = isValidRegistrationNumber(registrationNumber);
                if (!isValid) {
                    System.out.println("Invalid registration number format. Please enter again.");
                }
            } while (!isValid);

            // Iterate through each row of the sheet to search for the desired registration number
            for (Row row : sheet) {
                Cell regNumberCell = row.getCell(1); // Assuming registration number is in the second column (index 1)
                if (regNumberCell != null && regNumberCell.getCellType() == CellType.STRING) {
                    String regNumber = regNumberCell.getStringCellValue();
                    if (regNumber.equals(registrationNumber)) {
                        // Print the entire row if the registration number matches
                        for (Cell cell : row) {
                            System.out.print(cell.toString() + "\t");
                        }
                        System.out.println(); // Move to the next line after printing the row
                        return; // Exit the loop after finding the record
                    }
                }
            }
            // Print message if the registration number is not found
            System.out.println("Registration number not found.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // Method to validate the format and type of the registration number
    private static boolean isValidRegistrationNumber(String regNumber) {
        // Define a regex pattern for the registration number format
        String regex = "FS\\d{3}"; // Assuming registration number consists of 3 digits
        return regNumber.matches(regex);
    }
}

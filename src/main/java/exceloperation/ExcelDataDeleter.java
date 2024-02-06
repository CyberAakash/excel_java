package exceloperation;

import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

public class ExcelDataDeleter {
    public static void main(String[] args) {
        // Path to the Excel file
        String filePath = "example.xlsx";

        try (Scanner scanner = new Scanner(System.in);
             FileInputStream fis = new FileInputStream(new File(filePath));
             Workbook workbook = WorkbookFactory.create(fis)) {

            // Get the first sheet of the workbook
            Sheet sheet = workbook.getSheetAt(0); // Assuming the data is in the first sheet

            // Prompt the user to enter the registration number to delete
            System.out.print("Enter Registration Number to delete: ");
            String registrationNumber = scanner.nextLine().trim(); // Trim spaces from the input

            boolean found = false; // Flag to track if any matching registration number is found

            // Iterate through each row of the sheet
            for (int i = sheet.getLastRowNum(); i >= 1; i--) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    Cell regNumberCell = row.getCell(1); // Assuming registration number is in the second column (index 1)
                    if (regNumberCell != null && regNumberCell.getCellType() == CellType.STRING) {
                        String regNumber = regNumberCell.getStringCellValue().trim(); // Trim spaces from the cell value
                        if (regNumber.equals(registrationNumber)) {
                            // If registration number is found, delete the entire row
                            sheet.removeRow(row);
                            found = true;
                        }
                    }
                }
            }

            if (!found) {
                System.out.println("Registration Number not found. Please try again.");
            } else {
                System.out.println("Record(s) with Registration Number '" + registrationNumber + "' deleted successfully.");
            }

            // Write the changes back to the Excel file
            try (FileOutputStream fos = new FileOutputStream(filePath)) {
                workbook.write(fos);
            } catch (IOException e) {
                e.printStackTrace();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

package exceloperation;

import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

public class ExcelDataWriter {
    public static void main(String[] args) {
        // Path to the Excel file
        String filePath = "example2.xlsx";

        Scanner scanner = new Scanner(System.in);

        try {
            // Create a FileInputStream to read the Excel file
            FileInputStream fis = new FileInputStream(new File(filePath));

            // Create a Workbook object for the Excel file
            Workbook workbook = WorkbookFactory.create(fis);

            // Get the first sheet of the workbook
            Sheet sheet = workbook.getSheetAt(0); // Assuming the data is in the first sheet

            // Get the index of the last row with data
            int lastRowIndex = sheet.getLastRowNum();

            // Prompt the user for the number of records to add
            System.out.print("Enter the number of records to add: ");
            int numRecords = scanner.nextInt();
            scanner.nextLine(); // Consume newline character

            // Infer the starting registration number
            String lastRegNumber = "";
            if (lastRowIndex > 0) {
                Row lastRow = sheet.getRow(lastRowIndex);
                if (lastRow != null) {
                    Cell regNumberCell = lastRow.getCell(1); // Assuming registration number is in the second column (index 1)
                    if (regNumberCell != null) {
                        lastRegNumber = regNumberCell.getStringCellValue();
                    }
                }
            }

            // Input data for each record from the user
            for (int i = 0; i < numRecords; i++) {
                // Create a new row at the end of the sheet
                Row newRow = sheet.createRow(lastRowIndex + i + 1);

                // Prompt the user for other details
                System.out.println("Enter details for record " + (i + 1) + ":");
                System.out.print("Name: ");
                String name = scanner.nextLine();

                // Print the registration number based on the last registration number
                String newRegNumber = "FS" + String.format("%03d", Integer.parseInt(lastRegNumber.substring(2)) + i + 1);
                System.out.println("Generated Registration Number: " + newRegNumber);

                // Set the registration number for the new record
                newRow.createCell(1).setCellValue(newRegNumber);

                System.out.print("Score: ");
                double score = scanner.nextDouble();
                scanner.nextLine(); // Consume newline character

                // Populate the new row with user input
                newRow.createCell(0).setCellValue(name);
                newRow.createCell(2).setCellValue(score);
            }


            // Write the changes back to the Excel file
            try (FileOutputStream fos = new FileOutputStream(filePath)) {
                workbook.write(fos);
                System.out.println("New data written to the end of the Excel file with inferred data types.");
            } catch (IOException e) {
                e.printStackTrace();
            }

            // Close the workbook and FileInputStream
            workbook.close();
            fis.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

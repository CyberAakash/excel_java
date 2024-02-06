package exceloperation;

import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.Scanner;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ExcelRecordUpdater {
    public static void main(String[] args) {
        // Path to the Excel file
        String filePath = "output.xlsx";

        // Prompt the user to enter the registration number for the record they want to update
        Scanner scanner = new Scanner(System.in);
        String registrationNumber; // No need to initialize here
        Pattern regNumberPattern = Pattern.compile("^FS\\d{3}$");
        Matcher matcher;

        boolean exit = false;

        while (!exit) {
            System.out.print("Enter Registration Number to update (Type 'exit' to quit): ");
            registrationNumber = scanner.nextLine();

            if (registrationNumber.equalsIgnoreCase("exit")) {
                System.out.println("Exiting program...");
                return;
            }

            matcher = regNumberPattern.matcher(registrationNumber);
            if (!matcher.matches()) {
                System.out.println("Invalid registration number format. Please enter a valid registration number in the format FS###.");
                continue;
            }

            try {
                // Create a FileInputStream to read the Excel file
                FileInputStream fis = new FileInputStream(new File(filePath));

                // Create a Workbook object for the Excel file
                Workbook workbook = WorkbookFactory.create(fis);

                // Get the first sheet of the workbook
                Sheet sheet = workbook.getSheetAt(0); // Assuming the data is in the first sheet

                // Iterate through each row of the sheet
                boolean recordFound = false;
                for (Row row : sheet) {
                    // Skip the header row
                    if (row.getRowNum() == 0) {
                        continue;
                    }

                    // Get the registration number cell value from the row
                    Cell regNumberCell = row.getCell(1); // Assuming registration number is in the second column (index 1)
                    if (regNumberCell != null && regNumberCell.getCellType() == CellType.STRING) {
                        String regNumber = regNumberCell.getStringCellValue();

                        // Check if the registration number matches the one to update
                        if (regNumber.equals(registrationNumber)) {
                            // Prompt the user to enter the updated details
                            System.out.println("Enter updated details for registration number " + registrationNumber + ":");

                            // Validate and update name
                            System.out.print("Name: ");
                            String name = scanner.nextLine();
                            row.getCell(0).setCellValue(name);

                            // Validate and update score
                            double score;
                            while (true) {
                                System.out.print("Score: ");
                                if (scanner.hasNextDouble()) {
                                    score = scanner.nextDouble();
                                    scanner.nextLine(); // Consume newline character
                                    row.getCell(2).setCellValue(score);
                                    break;
                                } else {
                                    System.out.println("Invalid input. Please enter a valid score.");
                                    scanner.nextLine(); // Consume invalid input
                                }
                            }

                            System.out.println("Record updated successfully.");
                            recordFound = true;
                            break; // Exit the loop after updating the record
                        }
                    }
                }

                // If the record was not found, display an error message
                if (!recordFound) {
                    System.out.println("Registration number not found. Please try again.");
                }

                // Write the changes back to the Excel file
                try (FileOutputStream fos = new FileOutputStream(filePath)) {
                    workbook.write(fos);
                }

                // Close the workbook and FileInputStream
                workbook.close();
                fis.close();

            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}

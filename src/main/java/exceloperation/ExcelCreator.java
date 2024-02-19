package exceloperation;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashSet;
import java.util.Scanner;
import java.util.Set;

public class ExcelCreator {
    public static void main(String[] args) {
        try (Scanner scanner = new Scanner(System.in);
             Workbook workbook = new XSSFWorkbook()) {
            // Create a sheet named "Sheet1"
            Sheet sheet = workbook.createSheet("Sheet1");

            // Create header row
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Name");
            headerRow.createCell(1).setCellValue("Reg No");
            headerRow.createCell(2).setCellValue("Score");

            // Set to store unique registration numbers
            Set<String> regNumbers = new HashSet<>();

            // Prompt the user for the number of records
            System.out.print("Enter the number of records: ");
            int numRecords = scanner.nextInt();
            scanner.nextLine(); // Consume newline character

            // Input data from the user for each record
            for (int i = 0; i < numRecords; i++) {
                // Generate registration number in format "FS001", "FS002", ...
                String regNumber = String.format("FS%03d", i + 1);
                while (regNumbers.contains(regNumber)) {
                    // If generated registration number already exists, increment the counter
                    i++;
                    regNumber = String.format("FS%03d", i + 1);
                }
                regNumbers.add(regNumber);

                System.out.println("Enter details for record " + (i + 1) + ":");
                System.out.print("Name: ");
                String name = scanner.nextLine();
                System.out.println("Register Number: " + regNumber);
                System.out.print("Score: ");
                double score = scanner.nextDouble();
                scanner.nextLine(); // Consume newline character

                // Create a new row and populate it with user input
                Row row = sheet.createRow(i + 1);
                row.createCell(0).setCellValue(name);
                row.createCell(1).setCellValue(regNumber);
                row.createCell(2).setCellValue(score);
            }

            // Write the workbook to a file
            try (FileOutputStream fos = new FileOutputStream("example.xlsx")) {
                workbook.write(fos);
                System.out.println("Excel file created successfully!");
            } catch (IOException e) {
                e.printStackTrace();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}


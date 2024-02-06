package exceloperation;

import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.*;

public class ExcelDuplicateFinder {
    public static void main(String[] args) {
        // Path to the Excel file
        String filePath = "output.xlsx";

        // Map to store registration numbers and their occurrence counts
        Map<String, List<String>> regNumberRecordMap = new HashMap<>();

        try {
            // Create a FileInputStream to read the Excel file
            FileInputStream fis = new FileInputStream(new File(filePath));

            // Create a Workbook object for the Excel file
            Workbook workbook = WorkbookFactory.create(fis);

            // Get the first sheet of the workbook
            Sheet sheet = workbook.getSheetAt(0); // Assuming the data is in the first sheet

            // Iterate through each row of the sheet
            for (Row row : sheet) {
                // Skip the header row
                if (row.getRowNum() == 0) {
                    continue;
                }

                // Get the registration number and details from the row
                Cell regNumberCell = row.getCell(1); // Assuming registration number is in the second column (index 1)
                Cell nameCell = row.getCell(0); // Assuming name is in the first column (index 0)
                Cell scoreCell = row.getCell(2); // Assuming score is in the third column (index 2)

                if (regNumberCell != null && regNumberCell.getCellType() == CellType.STRING) {
                    String regNumber = regNumberCell.getStringCellValue();

                    // Update the record list associated with the registration number in the map
                    List<String> records = regNumberRecordMap.getOrDefault(regNumber, new ArrayList<>());
                    records.add(nameCell.getStringCellValue() + "\t" + regNumber + "\t" + scoreCell.toString());
                    regNumberRecordMap.put(regNumber, records);
                }
            }

            // Close the workbook and FileInputStream
            workbook.close();
            fis.close();

            // Display duplicate records details along with the number of duplicates
            boolean duplicatesFound = false;
            for (Map.Entry<String, List<String>> entry : regNumberRecordMap.entrySet()) {
                if (entry.getValue().size() > 1) {
                    System.out.println("Registration Number: " + entry.getKey() + " - " + (entry.getValue().size() - 1) + " duplicates");
                    // Display duplicate records
                    for (String record : entry.getValue()) {
                        System.out.println(record);
                    }
                    duplicatesFound = true;
                }
            }

            if (!duplicatesFound) {
                System.out.println("No duplicate records found.");
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

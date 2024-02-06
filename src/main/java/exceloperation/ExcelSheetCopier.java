package exceloperation;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
public class ExcelSheetCopier {
    public static void main(String[] args) {
        // Source and destination file paths
        String sourceFilePath = "example.xlsx";
        String destinationDirectoryPath = "destination";
        String destinationFilePath = "destination/example_copy.xlsx";

        try {
            // Create the destination directory if it doesn't exist
            File destinationDirectory = new File(destinationDirectoryPath);
            if (!destinationDirectory.exists()) {
                destinationDirectory.mkdirs();
            }

            // Create a FileInputStream to read the source Excel file
            FileInputStream fis = new FileInputStream(new File(sourceFilePath));

            // Create a Workbook object for the source Excel file
            Workbook sourceWorkbook = WorkbookFactory.create(fis);

            // Get the first sheet of the source workbook
            Sheet sourceSheet = sourceWorkbook.getSheetAt(0); // Assuming the data is in the first sheet

            // Create a new Workbook object for the destination Excel file
            Workbook destinationWorkbook = new XSSFWorkbook();

            // Create a new sheet in the destination workbook
            Sheet destinationSheet = destinationWorkbook.createSheet("Sheet1");

            // Copy data from source sheet to destination sheet
            for (int rowIndex = 0; rowIndex <= sourceSheet.getLastRowNum(); rowIndex++) {
                Row sourceRow = sourceSheet.getRow(rowIndex);
                Row destinationRow = destinationSheet.createRow(rowIndex);

                if (sourceRow != null) {
                    for (int cellIndex = 0; cellIndex < sourceRow.getLastCellNum(); cellIndex++) {
                        Cell sourceCell = sourceRow.getCell(cellIndex);
                        Cell destinationCell = destinationRow.createCell(cellIndex);

                        if (sourceCell != null) {
                            switch (sourceCell.getCellType()) {
                                case NUMERIC:
                                    destinationCell.setCellValue(sourceCell.getNumericCellValue());
                                    break;
                                case BOOLEAN:
                                    destinationCell.setCellValue(sourceCell.getBooleanCellValue());
                                    break;
                                case STRING:
                                    destinationCell.setCellValue(sourceCell.getStringCellValue());
                                    break;
                                // Handle other cell types if needed
                            }
                        }
                    }
                }
            }

            // Write the destination workbook to a new Excel file
            try (FileOutputStream fos = new FileOutputStream(destinationFilePath)) {
                destinationWorkbook.write(fos);
                System.out.println("Excel sheet copied successfully from '" + sourceFilePath + "' to '" + destinationFilePath + "'.");
            } catch (IOException e) {
                e.printStackTrace();
            }

            // Close the source and destination workbooks and FileInputStream
            sourceWorkbook.close();
            destinationWorkbook.close();
            fis.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
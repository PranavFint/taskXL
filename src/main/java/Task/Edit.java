package Task;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Edit {
    public static void main(String[] args) {
        Scanner scanner = new Scanner(System.in);

        // Ask for the Excel file path
        System.out.print("Enter the Excel file path: ");
        String filePath = scanner.nextLine();

        // Ask for the sheet number
        System.out.print("Enter the sheet number: ");
        int sheetNumber = scanner.nextInt();

        // Ask for the row number
        System.out.print("Enter the row number: ");
        int rowNumber = scanner.nextInt();

        // Ask for the column number
        System.out.print("Enter the column number: ");
        int columnNumber = scanner.nextInt();
        scanner.nextLine(); // Consume the newline character left by nextInt()

        // Ask for the new string to replace
        System.out.print("Enter the new string: ");
        String newString = scanner.nextLine();

        scanner.close();

        try {
            // Load the Excel file
            FileInputStream file = new FileInputStream(new File(filePath));
            Workbook workbook = new XSSFWorkbook(file);

            // Get the specified sheet
            Sheet sheet = workbook.getSheetAt(sheetNumber);

            // Get the specified row
            Row row = sheet.getRow(rowNumber);

            // Get the specified cell
            Cell cell = row.getCell(columnNumber);

            // Set the new value for the cell
            cell.setCellValue(newString);

            // Write the changes back to the Excel file
            FileOutputStream out = new FileOutputStream(new File(filePath));
            workbook.write(out);
            out.close();

            System.out.println("Excel file updated successfully!");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

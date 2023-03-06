package Task;

import java.io.BufferedReader;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.Scanner;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CsvFile {

    public static void main(String[] args) throws IOException {
        Scanner scanner = new Scanner(System.in);

        System.out.print("Enter the path of the CSV file: ");
        String csvFilePath = scanner.nextLine();

        scanner.close();

        String outputFilePath = "output.xlsx";
        String sheetName = "Sheet1";

        BufferedReader reader = new BufferedReader(new FileReader(csvFilePath));
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet(sheetName);

        String line = reader.readLine();
        int rowNumber = 0;

        while (line != null) {
            Row row = sheet.createRow(rowNumber++);
            String[] tokens = line.split(",");
            int columnNumber = 0;
            for (String token : tokens) {
                Cell cell = row.createCell(columnNumber++);
                cell.setCellValue(token);
            }
            line = reader.readLine();
        }

        reader.close();
        FileOutputStream outputStream = new FileOutputStream(outputFilePath);
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();
        System.out.println("Excel file created successfully.");
    }
}

package Task;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Edit{
    public static void main(String[] args) throws Exception {
        String filePath = "C:\\Users\\PranavKrishnamurthyB\\IdeaProjects\\readwrite\\gfgcontribute.xlsx";

        // Open the Excel file
        Workbook workbook = new XSSFWorkbook(new FileInputStream(filePath));

        // Get the first sheet from the workbook
        Sheet sheet = workbook.getSheetAt(0);

        // Get the fourth row from the sheet
        Row row = sheet.getRow(3);

        // Get the cell B in the fourth row
        Cell cell = row.getCell(1);

        // Set the value of the cell to "SuryaKumar"
        cell.setCellValue("SuryaKumar");

        // Save the changes to the Excel file
        FileOutputStream outputStream = new FileOutputStream(filePath);
        workbook.write(outputStream);
        outputStream.close();

        // Close the workbook
        workbook.close();
    }
}




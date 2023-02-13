package Task;

import jdk.internal.module.ModuleInfoExtender;
import org.apache.commons.compress.archivers.dump.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.*;

public class readfromfile {

    private static final String FILE_NAME
            = "C:\\Users\\PranavKrishnamurthyB\\Documents\\exel file\\dataset_2.xlsx";


    public static void main(String[] args) throws IOException {
        write();
    }

    // Method

        public static void write () throws IOException, InvalidFormatException
        {

            InputStream inp = new FileInputStream(FILE_NAME);
            Workbook wb = WorkbookFactory.create(inp);
            Sheet sheet = wb.getSheetAt(0);
            int num = sheet.getLastRowNum();
            Row row = sheet.createRow(++num);
            row.createCell(0).setCellValue("xyz");

        }
        ;

        {
            FileOutputStream fileOut = null;
            try {
                fileOut = new FileOutputStream(FILE_NAME);
            } catch (FileNotFoundException e) {
                throw new RuntimeException(e);
            }
            ModuleInfoExtender wb = null;
            try {
                wb.write(fileOut);
            } catch (IOException e) {
                throw new RuntimeException(e);
            }

            // Closing the file connections
            try {
                fileOut.close();
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        }

}

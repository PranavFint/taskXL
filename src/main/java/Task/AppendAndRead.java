package Task;
import org.apache.commons.compress.archivers.dump.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

public class AppendAndRead {
    private static final String FILE_NAME
            = "C:\\Users\\PranavKrishnamurthyB\\Documents\\exel file\\dataset_2.xlsx";
    public static void write() throws IOException, InvalidFormatException {

        InputStream inp = new FileInputStream(FILE_NAME);
        Workbook wb = WorkbookFactory.create(inp);
        Sheet sheet = wb.getSheetAt(0);
    }

}

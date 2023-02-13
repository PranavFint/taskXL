package Task;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Properties;
import javax.mail.Folder;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.NoSuchProviderException;
import javax.mail.Session;
import javax.mail.Store;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadEmailAndUpdateExcel {
    public static void main(String[] args) {
        String email = "kbpranav@ymail.com";
        String password = "krishna01";
        String otp = readOTPFromEmail(email, password);
        updateOTPInExcel(otp);
    }

    private static String readOTPFromEmail(String email, String password) {
        String otp = "";
        Properties properties = new Properties();
        properties.put("mail.pop3.host", "pop.gmail.com");
        properties.put("mail.pop3.port", "995");
        properties.put("mail.pop3.starttls.enable", "true");
        Session emailSession = Session.getDefaultInstance(properties);
        try {
            Store store = emailSession.getStore("pop3s");
            store.connect(email, password);
            Folder emailFolder = store.getFolder("INBOX");
            emailFolder.open(Folder.READ_ONLY);
            Message[] messages = emailFolder.getMessages();
            for (int i = 0; i < messages.length; i++) {
                Message message = messages[i];
                String subject = message.getSubject();
                if (subject.equals("OTP")) {
                    otp = message.getContent().toString();
                    break;
                }
            }
            emailFolder.close(false);
            store.close();
        } catch (NoSuchProviderException e) {
            e.printStackTrace();
        } catch (MessagingException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return otp;
    }

    private static void updateOTPInExcel(String otp) {
        try {
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("OTP");
            Row row = sheet.createRow(0);
            Cell cell = row.createCell(0);
            cell.setCellValue(otp);
            workbook.write(new FileOutputStream("otp.xlsx"));
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}


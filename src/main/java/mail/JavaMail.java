package mail;

// Java Mail 相關
import javax.mail.*;
import javax.mail.internet.*;
import javax.mail.Authenticator;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.*;

// Java Activation Framework 相關
import javax.activation.*;
import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;

// Java IO 相關
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

// Java Collections 相關
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Properties;

// Apache POI 相關
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

// INI 設定檔相關
import org.ini4j.Ini;

public class JavaMail {
    // CONFIG PATHS
    private static final String CONFIG_FILE = "config.ini";
    private static final String CONTACT_EXCEL = "data/0419 聯絡資料.xlsx";
    private static final String CERT_DIR = "data/0419 證書";
    private static final String COURSE_NAME = "2025 未來造浪 AI Studio";

    public static void main(String[] args) {
        try {
            // 1. 讀取 INI 設定
            Ini ini = new Ini(new File(CONFIG_FILE));
            Ini.Section smtpSec = ini.get("SMTP");
            Ini.Section testSec = ini.get("TEST");

            String host = smtpSec.get("server");
            int port = Integer.parseInt(smtpSec.get("port"));
            String user = smtpSec.get("username");
            String pass = smtpSec.get("password");
            String senderEmail = smtpSec.get("sender_email");
            boolean useTls = Boolean.parseBoolean(smtpSec.get("use_tls"));

            boolean testMode = false;
            String testEmail = null;
            if (testSec != null) {
                testMode = Boolean.parseBoolean(testSec.get("enable_test_mode"));
                testEmail = testSec.get("recipient_email");
            }

            // 2. 構建 Session
            Properties props = new Properties();
            props.put("mail.smtp.auth", "true");
            props.put("mail.smtp.host", host);
            props.put("mail.smtp.port", String.valueOf(port));
            if (useTls) {
                props.put("mail.smtp.starttls.enable", "true");
            } else {
                props.put("mail.smtp.ssl.enable", "true");
            }
            Session session = Session.getInstance(props, new Auth(user, pass));
            session.setDebug(true);

            // 3. 載入聯絡人資料
            Map<String, String> contacts = loadContacts(CONTACT_EXCEL);

            // 4. 載入證書文件映射
            Map<String, Path> certMap = loadCertificates(CERT_DIR);

            int success = 0, fail = 0, skip = 0;
            for (Map.Entry<String, String> entry : contacts.entrySet()) {
                String name = entry.getKey();
                String email = entry.getValue();
                if (name.isEmpty() || email.isEmpty()) {
                    skip++; continue;
                }
                Path cert = certMap.get(name);
                if (cert == null) {
                    skip++; continue;
                }
                String to = testMode ? testEmail : email;
                String subject = String.format("「%s」課程證書寄發通知", COURSE_NAME)
                        + (testMode ? " (測試模式)" : "");
                String body = buildBody(name, testMode, testEmail);
                try {
                    send(session, senderEmail, to, subject, body, cert);
                    success++;
                } catch (Exception e) {
                    fail++;
                    e.printStackTrace();
                }
                // wait interval
                Thread.sleep(2000);
            }

            System.out.printf("完成：成功 %d / 失敗 %d / 跳過 %d\n", success, fail, skip);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static Map<String, String> loadContacts(String excelPath) throws IOException {
        Map<String, String> map = new HashMap<>();
        try (XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(excelPath))) {
            Sheet sheet = wb.getSheetAt(0);
            Iterator<Row> rows = sheet.rowIterator();
            Row header = rows.next();
            int idxName=-1, idxEmail=-1;
            for (Cell c : header) {
                String v = c.getStringCellValue();
                if ("姓名".equals(v)) idxName = c.getColumnIndex();
                if ("電子郵件".equals(v)) idxEmail = c.getColumnIndex();
            }
            while (rows.hasNext()) {
                Row r = rows.next();
                String name = r.getCell(idxName).getStringCellValue().trim();
                String email = r.getCell(idxEmail).getStringCellValue().trim();
                map.put(name, email);
            }
        }
        return map;
    }

    private static Map<String, Path> loadCertificates(String dir) throws IOException {
        Map<String, Path> map = new HashMap<>();
        Files.list(Paths.get(dir))
            .filter(p -> p.toString().endsWith(".pdf"))
            .forEach(p -> {
                String file = p.getFileName().toString();
                String stem = file.substring(0, file.lastIndexOf('.'));
                String name = stem.substring(stem.lastIndexOf('-')+1).trim();
                map.put(name, p);
            });
        return map;
    }

    private static String buildBody(String name, boolean testMode, String testEmail) {
        String prefix = name + " 同學，您好：\n\n";
        if (testMode) {
            prefix += String.format("(此為測試模式郵件，實際寄送至 %s)\n\n", testEmail);
        }
        return prefix +
               "感謝您參加「" + COURSE_NAME + "」課程... [郵件內容略]";
    }

    private static void send(Session session, String from, String to,
                             String subject, String body, Path attachment)
            throws MessagingException {
        MimeMessage msg = new MimeMessage(session);
        msg.setFrom(new InternetAddress(from));
        msg.setRecipient(Message.RecipientType.TO, new InternetAddress(to));
        msg.setSubject(subject);

        MimeBodyPart text = new MimeBodyPart();
        text.setText(body, "UTF-8");

        MimeBodyPart filePart = new MimeBodyPart();
        DataSource source = new FileDataSource(attachment.toFile());
        filePart.setDataHandler(new DataHandler(source));
        filePart.setFileName(attachment.getFileName().toString());

        MimeMultipart mp = new MimeMultipart();
        mp.addBodyPart(text);
        mp.addBodyPart(filePart);
        msg.setContent(mp);

        Transport.send(msg);
    }
}

class Auth extends Authenticator {
    private final String user, pass;
    public Auth(String user, String pass) { this.user=user; this.pass=pass; }
    @Override
    protected PasswordAuthentication getPasswordAuthentication() {
        return new PasswordAuthentication(user, pass);
    }
}

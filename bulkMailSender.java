import com.google.common.collect.ArrayListMultimap;
import com.google.common.collect.Multimap;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.mail.*;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;
import java.io.*;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

public class bulkMailSender {

    private static final Properties PROPERTIES = new Properties();
    private static final String USERNAME = "";   //change it
    private static final String PASSWORD = "rdou dfdf opws voxp";   //change it
    private static final String HOST = "smtp.gmail.com";
    static Multimap<String, LocalDateTime> emailsSentTime = ArrayListMultimap.create();

    public static void main(String[] args) throws IOException, NoSuchProviderException {

        List<String> ls = new ArrayList<>();
        try (InputStream inputStream = new FileInputStream("emails.xlsx")) {
            Workbook workbook = WorkbookFactory.create(inputStream);
            Sheet sheet = workbook.getSheetAt(0); // Assuming you want to read from the first sheet

            for (Row row : sheet) {
                Cell cell = row.getCell(1);
                if (cell != null) {
                    switch (cell.getCellType()) {
                        case STRING:
                            ls.add(cell.getStringCellValue());
                            break;

                        default:
                            System.out.println("Unknown cell type");
                    }
                } else {
                    System.out.println("CELL IS NULL");
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        System.out.println(ls.size());
        boolean status = sendmailsbulk(ls);

//        for(String mail : ls) {
//
//
//            if (status) {
//                updateStatusInExcel("emails.xlsx", mail);
//            } else {
//                updateStatusInExcel("emails.xlsx", mail);
//            }
//        }


    }

//    private static void updateStatusInExcel(String excelFilePath, String email) {
//        try (InputStream inputStream = new FileInputStream(excelFilePath)) {
//            Workbook workbook = new XSSFWorkbook(inputStream);
//            Sheet sheet = workbook.getSheetAt(0); // Assuming the recipients are in the first sheet
//            for (Row row : sheet) {
//                Cell cell = row.getCell(1); // Assuming email addresses are in the first column
//                if (cell.getStringCellValue().equals(email)) {
//                    Cell statusCell = row.createCell(2); // Assuming status will be written in the second column
//                    statusCell.setCellValue(status);
//                    Cell timestampCell = row.createCell(3); // Assuming timestamp will be written in the third column
//                    timestampCell.setCellValue(timestamp.format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss")));
//                    break; // Exit loop after updating status and timestamp
//                }
//            }
//            // Write the changes back to the Excel file
//            try (OutputStream outputStream = new FileOutputStream(excelFilePath)) {
//                workbook.write(outputStream);
//            }
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
//    }

    public static boolean sendmailsbulk(List<String> ls) throws NoSuchProviderException {
        Properties props = new Properties();
        props.put("mail.smtp.host", "smtp.gmail.com");
        props.put("mail.smtp.port", "587");
        props.put("mail.smtp.auth", "true");
        props.put("mail.smtp.starttls.enable", "true");

        Session session = Session.getInstance(props,
                new javax.mail.Authenticator() {
                    protected PasswordAuthentication getPasswordAuthentication() {
                        return new PasswordAuthentication(USERNAME, PASSWORD);
                    }
                });
        LocalDateTime sentTime = LocalDateTime.now();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
        sentTime.format(formatter);
        int rowNum = 0; // Start from the first row
        try (InputStream inputStream = new FileInputStream("emails.xlsx")) {
            Workbook workbook = WorkbookFactory.create(inputStream);
            Sheet sheet = workbook.getSheetAt(0);
            try {
                for (String recipient : ls) {
                    Row row = sheet.getRow(rowNum);
                    if (row == null) {
                        row = sheet.createRow(rowNum);
                    }

                    Message message = new MimeMessage(session);
                    message.setFrom(new InternetAddress(USERNAME));
                    message.setRecipients(Message.RecipientType.TO,
                            InternetAddress.parse(recipient));
                    message.setSubject("Your Subject Here");
                    message.setText("Dear recipient," +
                            "\n\n This is the content of your email!");
                    try {
                        Transport.send(message);
                    //    emailsSentTime.put("sent", sentTime);
                        row.createCell(0).setCellValue(recipient); // Recipient
                        row.createCell(1).setCellValue(sentTime.format(formatter)); // Timestamp
                        row.createCell(2).setCellValue("Sent");
                    } catch (MessagingException e) {
                       // emailsSentTime.put("not sent", sentTime);
                        row.createCell(0).setCellValue(recipient); // Recipient
                        row.createCell(1).setCellValue(sentTime.format(formatter)); // Timestamp
                        row.createCell(2).setCellValue("Failed"); // Status

                    }
                    System.out.println("Email sent to: " + recipient);
                    rowNum++;
                    Thread.sleep(1000); // Adding a delay to prevent rate limiting
                    if(rowNum==6){
                        break;
                    }
                }

                System.out.println("All emails sent successfully!");
                try (OutputStream outputStream = new FileOutputStream("emails.xlsx")) {
                    workbook.write(outputStream);
                }
            } catch (MessagingException | InterruptedException e) {
                emailsSentTime.put("not sent", sentTime);

                e.printStackTrace();
                return false;
            }
            return true;
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}

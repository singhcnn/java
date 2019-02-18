package com.mail.data.MailApiManager;

import java.io.BufferedReader;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Properties;

import javax.mail.Address;
import javax.mail.Folder;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.NoSuchProviderException;
import javax.mail.Part;
import javax.mail.Session;
import javax.mail.Store;
import javax.mail.internet.MimeBodyPart;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class App {
	private static final String HOST = "pop.gmail.com";
	private static final String USER = "singh.cnn@gmail.com";
	private static final String MAILSTORTYPE = "pop3";
	private static final String PASSWD = "v6000sandy";
	private static Folder emailFolder = null;
	private static final String[] columns = { "FROM", "SUBJECT", "DATE ", "TIME", "LABEL",
			"ATTACHMENT NAME/FILE NAME" };
	private static final String PATTERN = "yyyy-MM-dd";

	/*
	 * This method would print FROM,TO and SUBJECT of the message
	 */

	public static void writeEnvelope(Message m, Row row) throws Exception {

		Address[] a;

		// FROM
		if ((a = m.getFrom()) != null) {
			for (int j = 0; j < a.length; j++)
				row.createCell(0).setCellValue(a[j].toString());
		}
		// SUBJECT
		if (m.getSubject() != null) {
			row.createCell(1).setCellValue(m.getSubject());
		}
		// Date
		if (m.getSentDate() != null) {
			SimpleDateFormat simpleDateFormat = new SimpleDateFormat(PATTERN);
			String date = simpleDateFormat.format(m.getSentDate());
			row.createCell(2).setCellValue(date);
		}
		// Time
		if (m.getSentDate().getTime() != 0) {
			final long timestamp = m.getSentDate().getTime();
			final Calendar cal = Calendar.getInstance();
			cal.setTimeInMillis(timestamp);
			final String timeString = new SimpleDateFormat("HH:mm:ss:SSS").format(cal.getTime());
			row.createCell(3).setCellValue(timeString);
		}
		// Label
		if (emailFolder.getName() != null) {
			row.createCell(4).setCellValue(emailFolder.getName());
		}
	}

	public static void main(String[] args) throws IOException {

		App.fetch(HOST, MAILSTORTYPE, USER, PASSWD);
	}

	public static void fetch(String pop3Host, String storeType, String user, String password) {
		try {
			// create properties field
			Properties properties = new Properties();
			properties.put("mail.store.protocol", "pop3");
			properties.put("mail.pop3.host", pop3Host);
			properties.put("mail.pop3.port", "995");
			properties.put("mail.pop3.starttls.enable", "true");
			Session emailSession = Session.getDefaultInstance(properties);

			// create the POP3 store object and connect with the pop server
			Store store = emailSession.getStore("pop3s");

			store.connect(pop3Host, user, password);

			// create the folder object and open it
			emailFolder = store.getDefaultFolder().getFolder("INBOX");
			emailFolder.open(Folder.READ_ONLY);

			BufferedReader reader = new BufferedReader(new InputStreamReader(System.in));
			// retrieve the messages from the folder in an array and print it
			Message[] messages = emailFolder.getMessages(1, 100);
			System.out.println("messages.length---" + messages.length);

			for (int i = 0; i < messages.length; i++) {
				Message message = messages[i];
				System.out.println("---------------------------------");
				writePart(message);
				String line = reader.readLine();
				if ("YES".equals(line)) {
					message.writeTo(System.out);
				} else if ("QUIT".equals(line)) {
					break;
				}
			}

			// close the store and folder objects
			emailFolder.close(false);
			store.close();

		} catch (NoSuchProviderException e) {
			e.printStackTrace();
		} catch (MessagingException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/*
	 * This method checks for content-type based on which, it processes and fetch
	 * the data
	 */
	public static void writePart(Part p) throws Exception {
		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet("MailData");

		CreationHelper createHelper = workbook.getCreationHelper();
		// Create a Font for styling header cells
		Font headerFont = workbook.createFont();
		headerFont.setBold(true);
		headerFont.setFontHeightInPoints((short) 14);
		headerFont.setColor(IndexedColors.RED.getIndex());

		// Create a CellStyle with the font
		CellStyle headerCellStyle = workbook.createCellStyle();
		headerCellStyle.setFont(headerFont);
		// Create a Row
		Row headerRow = sheet.createRow(0);

		// Create cells
		for (int i = 0; i < columns.length; i++) {
			Cell cell = headerRow.createCell(i);
			cell.setCellValue(columns[i]);
			cell.setCellStyle(headerCellStyle);
		}
		// Create Cell Style for formatting Date

		CellStyle dateCellStyle = workbook.createCellStyle();
		dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd-MM-yyyy"));

		int rowNum = 1;
		Row row = sheet.createRow(rowNum++);

		if (p instanceof Message)
			writeEnvelope((Message) p, row);
		System.out.println("----------------------------");
		System.out.println("CONTENT-TYPE: " + p.getContentType());

		// check if the content has attachment
		if (p.isMimeType("multipart/*")) {

			System.out.println("This is a Multipart");
			System.out.println("---------------------------");
			Multipart mp = (Multipart) p.getContent();
			int count = mp.getCount();

			for (int i = 0; i < count; i++) {
				MimeBodyPart part = (MimeBodyPart) mp.getBodyPart(i);
				row.createCell(5).setCellValue(part.getFileName());
				writePart(part);
			}

		} else {
			System.out.println("This is an unknown type");
			System.out.println("---------------------------");
		}

		for (int i = 0; i < columns.length; i++) {
			sheet.autoSizeColumn(i);
		}
// Write the output to a file
		FileOutputStream fileOut = new FileOutputStream("poi-generated-file.xlsx");
		workbook.write(fileOut);
		fileOut.close();
		fileOut.flush();
// Closing the workbook
		workbook.close();

	}

}

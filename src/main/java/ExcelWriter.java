import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

public class ExcelWriter {

    private static String[] columns = {"Name","Surname", "Date Of Birth", "Points"};

    private static List<Team> team =  new ArrayList<>();
    private static Object InvalidFormatException;

    static {
        Calendar dateOfBirth = Calendar.getInstance();
        dateOfBirth.set(1985, 9, 21);
        team.add(new Team("Michael", "Brown",
                dateOfBirth.getTime(), 5434.0));

        dateOfBirth.set(1997, 6, 18);
        team.add(new Team("Tom", "Smith",
                dateOfBirth.getTime(), 7643.0));

        dateOfBirth.set(1997, 5, 8);
        team.add(new Team("Steve", "White",
                dateOfBirth.getTime(), 18000.0));
    }

    public static void main(String[] args) throws IOException, InvalidFormatException {

        // Create a Workbook
        Workbook workbook = new XSSFWorkbook();     // new HSSFWorkbook() for generating `.xls` file

        /* CreationHelper helps us create instances for various things like DataFormat,
           Hyperlink, RichTextString etc in a format (HSSF, XSSF) independent way */
        CreationHelper createHelper = workbook.getCreationHelper();

        // Create a Sheet
        Sheet sheet = workbook.createSheet("Team");

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

        // Creating cells
        for (int i = 0; i < columns.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(headerCellStyle);
        }

        // Cell Style for formatting Date
        CellStyle dateCellStyle = workbook.createCellStyle();
        dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd-MM-yyyy"));

        // Create Other rows and cells with employees data
        int rowNum = 1;
        for (Team team : team) {
            Row row = sheet.createRow(rowNum++);

            row.createCell(0)
                    .setCellValue(team.getName());

            row.createCell(1)
                    .setCellValue(team.getSurname());

            Cell dateOfBirthCell = row.createCell(2);
            dateOfBirthCell.setCellValue(team.getDateOfBirth());
            dateOfBirthCell.setCellStyle(dateCellStyle);

            row.createCell(3)
                    .setCellValue(team.getPoints());
        }

        // Resize all columns to fit the content size
        for (int i = 0; i < columns.length; i++) {
            sheet.autoSizeColumn(i);
        }

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("poi-generated-file.xlsx");
        workbook.write(fileOut);
        fileOut.close();

        workbook.close();
        modifyExistingWorkbook();
    }

    // Example to modify an existing excel file
    private static void modifyExistingWorkbook() throws InvalidFormatException, IOException {
        String excelFilePath = "existing-spreadsheet.xlsx";

        try {
            FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
            Workbook workbook = WorkbookFactory.create(inputStream);

            Sheet sheet = workbook.getSheetAt(0);

            Object[][] list = {
                    {"Michelangelo Buonarroti", "Florence Republic", "Italian Renaissance"},
                    {"Leonardo da Vinci", "Florence Republic", "Italian Renaissance"},
                    {"Rafael Santi", "Holy Roman Empire", "Italian Renaissance"},
                    {"Titian Vecellio", "Venetian Republic", "Italian Renaissance"},
                    {"Rembrandt", "Dutch Republic", "Dutch Golden Age Baroque"}
            };

            int rowCount = sheet.getLastRowNum();

            for (Object[] artist : list) {
                Row row = sheet.createRow(++rowCount);

                int columnCount = 0;

                Cell cell = row.createCell(columnCount);
                cell.setCellValue(rowCount);

                for (Object field : artist) {
                    cell = row.createCell(++columnCount);
                    if (field instanceof String) {
                        cell.setCellValue((String) field);
                    } else if (field instanceof Integer) {
                        cell.setCellValue((Integer) field);
                    }
                }

            }

            inputStream.close();

            FileOutputStream outputStream = new FileOutputStream("existing-spreadsheet.xlsx");
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();

        } catch (IOException | EncryptedDocumentException
                | InvalidFormatException ex) {
            ex.printStackTrace();
        }
        }
    }



class Team {
    private String name;

    private String surname;

    private Date dateOfBirth;

    private double points;

    public Team(String name, String surname, Date dateOfBirth, double points) {
        this.name = name;
        this.surname = surname;
        this.dateOfBirth = dateOfBirth;
        this.points = points;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getSurname() {
        return surname;
    }

    public void setSurname(String surname) {
        this.surname = surname;
    }

    public Date getDateOfBirth() {
        return dateOfBirth;
    }

    public void setDateOfBirth(Date dateOfBirth) {
        this.dateOfBirth = dateOfBirth;
    }

    public double getPoints() {
        return points;
    }

    public void setPoints(double points) {
        this.points = points;
    }
}

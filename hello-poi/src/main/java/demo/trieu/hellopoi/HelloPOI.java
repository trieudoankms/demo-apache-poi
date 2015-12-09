package demo.trieu.hellopoi;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Date;

/**
 * Created by trieudoan on 12/8/2015.
 */
public class HelloPOI {
    private static final String TEST_FILE_1 = "test.xlsx";
    private static final String TEST_FILE_2 = "test2.xlsx";

    public static void main(String[] args) {
        helloDemo();

        System.out.println("-------");
        demoReadingTestFile();

        System.out.println("-------");
        demoEditingTestFile();
    }

    private static void helloDemo() {
        System.out.println("Hello demo!");
        Workbook wb = new XSSFWorkbook();
        CreationHelper createHelper = wb.getCreationHelper();

        // Create sheets
        Sheet sheet1 = wb.createSheet("new sheet");
        wb.createSheet("second sheet");

        // Create a row and put some cells in it
        Row row = sheet1.createRow(0);
        // Create cells
        Cell cell = row.createCell(0);
        cell.setCellValue(1);
        row.createCell(1).setCellValue(1.2);
        row.createCell(2).setCellValue(createHelper.createRichTextString("This is a string!"));
        row.createCell(3).setCellValue(true);
        row.createCell(4).setCellValue(new Date());
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("mm/dd/yyyy"));
        Cell cell5 = row.createCell(5);
        cell5.setCellValue(new Date());
        cell5.setCellStyle(cellStyle);

        try {
            FileOutputStream fileOut = new FileOutputStream("workbook.xlsx");
            wb.write(fileOut);
            fileOut.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void demoReadingTestFile() {
        if(Files.notExists(Paths.get(TEST_FILE_1))) {
            System.out.println("Create " + TEST_FILE_1);
            createTestFile(TEST_FILE_1);
        }

        System.out.println("Read " + TEST_FILE_1);
        try {
            readFile(TEST_FILE_1);
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
    }

    private static void createTestFile(String name) {
        Workbook wb = new XSSFWorkbook();
        CreationHelper createHelper = wb.getCreationHelper();

        Sheet sheet = wb.createSheet("books");

        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue(createHelper.createRichTextString("ID"));
        header.createCell(1).setCellValue(createHelper.createRichTextString("Name"));
        header.createCell(2).setCellValue(createHelper.createRichTextString("Author"));
        header.createCell(3).setCellValue(createHelper.createRichTextString("Price"));

        Row row1 = sheet.createRow(1);
        row1.createCell(0).setCellValue(createHelper.createRichTextString("A00001"));
        row1.createCell(1).setCellValue(createHelper.createRichTextString("Head First Java, 2nd Edition"));
        row1.createCell(2).setCellValue(createHelper.createRichTextString("Kathy Sierra and Bert Bates"));
        row1.createCell(3).setCellValue(26.07);

        Row row2 = sheet.createRow(2);
        row2.createCell(0).setCellValue(createHelper.createRichTextString("A00002"));
        row2.createCell(1).setCellValue(createHelper.createRichTextString("Effective Java (2nd Edition)"));
        row2.createCell(2).setCellValue(createHelper.createRichTextString("Joshua Bloch"));
        row2.createCell(3).setCellValue(31.62);

        Row row3 = sheet.createRow(3);
        row3.createCell(0).setCellValue(createHelper.createRichTextString("A00003"));
        row3.createCell(1).setCellValue(createHelper.createRichTextString("Java: A Beginner's Guide, Sixth Edition"));
        row3.createCell(2).setCellValue(createHelper.createRichTextString("Herbert Schildt"));
        row3.createCell(3).setCellValue(23.35);

        Row row4 = sheet.createRow(4);
        row4.createCell(0).setCellValue(createHelper.createRichTextString("A00004"));
        row4.createCell(1).setCellValue(createHelper.createRichTextString("Elements of Programming Interviews in Java: The Insiders' Guide"));
        row4.createCell(2).setCellValue(createHelper.createRichTextString("Adnan Aziz and Tsung-Hsien Lee"));
        row4.createCell(3).setCellValue(28.79);

        try {
            FileOutputStream fileOut = new FileOutputStream(name);
            wb.write(fileOut);
            fileOut.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void readFile(String name) throws IOException, InvalidFormatException {
        Workbook wb = WorkbookFactory.create(new File(name));


        Sheet sheet1 = wb.getSheetAt(0);
        for (Row row : sheet1) {
            for (Cell cell : row) {
                /*CellReference cellRef = new CellReference(row.getRowNum(), cell.getColumnIndex());
                System.out.print(cellRef.formatAsString());
                System.out.print(" - ");*/

                switch (cell.getCellType()) {
                    case Cell.CELL_TYPE_STRING:
                        System.out.print(cell.getRichStringCellValue().getString());
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        if (DateUtil.isCellDateFormatted(cell)) {
                            System.out.print(cell.getDateCellValue());
                        } else {
                            System.out.print(cell.getNumericCellValue());
                        }
                        break;
                    case Cell.CELL_TYPE_BOOLEAN:
                        System.out.print(cell.getBooleanCellValue());
                        break;
                    case Cell.CELL_TYPE_FORMULA:
                        System.out.print(cell.getCellFormula());
                        break;
                    default:
                        System.out.print("\t");
                }
                System.out.print("\t");
            }
            System.out.println();
        }
    }

    private static void demoEditingTestFile() {
        System.out.println(String.format("Edit data in %s and save it as %s", TEST_FILE_1, TEST_FILE_2));
        try {
            editTestFile(TEST_FILE_1, TEST_FILE_2);
            System.out.println(String.format("The content of %s is: ", TEST_FILE_2));
            readFile(TEST_FILE_2);
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
    }

    public static void editTestFile(String inputFileName, String outputFileName) throws IOException, InvalidFormatException {
        Workbook wb = WorkbookFactory.create(new File(inputFileName));
        CreationHelper createHelper = wb.getCreationHelper();
        Sheet sheet = wb.getSheetAt(0);

        // Update price of book at the fourth row
        Row row4 = sheet.getRow(4);
        Cell priceCell = row4.getCell(3);
        priceCell.setCellValue(100.0f);

        // Create a new book
        Row newRow = sheet.createRow(sheet.getLastRowNum() + 1);
        newRow.createCell(0).setCellValue(createHelper.createRichTextString("A00005"));
        newRow.createCell(1).setCellValue(createHelper.createRichTextString("The C++ Programming Language (3rd Edition)"));
        newRow.createCell(2).setCellValue(createHelper.createRichTextString("Bjarne Stroustrup"));
        newRow.createCell(3).setCellValue(59.05);

        FileOutputStream fileOut = new FileOutputStream(outputFileName);
        wb.write(fileOut);
        fileOut.close();
    }
}

package demo.poi;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadWriteExcelFile {

    public static void readXLSFile() throws IOException {
        InputStream ExcelFileToRead = new FileInputStream("C:/Test.xls");
        HSSFWorkbook wb = new HSSFWorkbook(ExcelFileToRead);

        HSSFSheet sheet = wb.getSheetAt(0);
        HSSFRow row;
        HSSFCell cell;

        Iterator rows = sheet.rowIterator();

        while (rows.hasNext()) {
            row = (HSSFRow) rows.next();
            Iterator cells = row.cellIterator();

            while (cells.hasNext()) {
                cell = (HSSFCell) cells.next();

                if (cell.getCellType() == HSSFCell.CELL_TYPE_STRING) {
                    System.out.print(cell.getStringCellValue() + " ");
                } else if (cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
                    System.out.print(cell.getNumericCellValue() + " ");
                } else {
                    //U Can Handel Boolean, Formula, Errors
                }
            }
            System.out.println();
        }

    }

    public static void writeXLSFile() throws IOException {

        String excelFileName = "C:/Test.xls";//name of excel file

        String sheetName = "Sheet1";//name of sheet

        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.createSheet(sheetName);

        //iterating r number of rows
        for (int r = 0; r < 5; r++) {
            HSSFRow row = sheet.createRow(r);

            //iterating c number of columns
            for (int c = 0; c < 5; c++) {
                HSSFCell cell = row.createCell(c);

                cell.setCellValue("Cell " + r + " " + c);
            }
        }

        FileOutputStream fileOut = new FileOutputStream(excelFileName);

        //write this workbook to an Outputstream.
        wb.write(fileOut);
        fileOut.flush();
        fileOut.close();
    }

    public static void mergeXLSXFile() throws IOException {
        InputStream ExcelFileToRead = new FileInputStream("C:\\MergeExcel\\KDW Updated Data capture matrix V3 for DSV-Survey input_ Survey Return 150716 .xlsx");
        XSSFWorkbook wb = new XSSFWorkbook(ExcelFileToRead);
        XSSFSheet sheet = wb.getSheet("DCV data to WFMT");


        XSSFWorkbook wbWrite = new XSSFWorkbook();
        XSSFSheet sheetWrite = wbWrite.createSheet("MergedSheet");

        XSSFRow row;
        XSSFCell cell;
        int r = 0;

        Iterator rows = sheet.rowIterator();
        while (rows.hasNext()) {
            XSSFRow rowWrite = sheetWrite.createRow(r++);
            row = (XSSFRow) rows.next();
            Iterator cells = row.cellIterator();
            int c = 0;
            while (cells.hasNext()) {
                XSSFCell cellWrite = rowWrite.createCell(c);
                cell = (XSSFCell) cells.next();
                int cellType = cell.getCellType();
                if (cellType == XSSFCell.CELL_TYPE_STRING) {
                    cellWrite.setCellValue(cell.getStringCellValue());
                } else if (cellType == XSSFCell.CELL_TYPE_NUMERIC) {
                    cellWrite.setCellValue(cell.getNumericCellValue());
                } else if(cellType == XSSFCell.CELL_TYPE_BOOLEAN){
                    cellWrite.setCellValue(cell.getBooleanCellValue());
                } else if(cellType == XSSFCell.CELL_TYPE_BLANK){
                    cellWrite.setCellValue("");
                }
                c++;
            }
        }

        FileOutputStream fileOut = new FileOutputStream("C:\\MergeExcel\\test.xlsx");
        wbWrite.write(fileOut);
        fileOut.flush();
        fileOut.close();
    }

    public static void writeXLSXFile() throws IOException {

        String excelFileName = "C:\\MergeExcel\\test.xlsx";//name of excel file
        String sheetName = "Sheet1";//name of sheet

        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet(sheetName);

        //iterating r number of rows
        for (int r = 0; r < 5; r++) {
            XSSFRow row = sheet.createRow(r);

            //iterating c number of columns
            for (int c = 0; c < 5; c++) {
                XSSFCell cell = row.createCell(c);

                cell.setCellValue("Cell " + r + " " + c);
            }
        }

        FileOutputStream fileOut = new FileOutputStream(excelFileName);

        //write this workbook to an Outputstream.
        wb.write(fileOut);
        fileOut.flush();
        fileOut.close();
    }

    public static void main(String[] args) throws IOException {

       /* writeXLSFile();
        readXLSFile();*/

        mergeXLSXFile();
        System.out.println("Operation finished!!");
        // readXLSXFile();

    }

}
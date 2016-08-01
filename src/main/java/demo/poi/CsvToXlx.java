package demo.poi;

import org.apache.poi.xssf.usermodel.*;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.util.ArrayList;
import java.util.List;

public class CsvToXlx {

    public static void csvToXLSX(XSSFWorkbook workBook, String filename, String sheetno) {
        try {
            String csvFile = filename; //csv file address
            String excelFile = "C:\\MergeExcel\\merge.xlsx"; //xlsx file address
            //XSSFWorkbook workBook = new XSSFWorkbook();
            XSSFSheet sheet = workBook.createSheet(sheetno);
            String currentLine = null;
            int RowNum = 0;
            BufferedReader br = new BufferedReader(new FileReader(csvFile));
            while ((currentLine = br.readLine()) != null) {
                try {
                    String str[] = currentLine.split(",");
                    XSSFRow currentRow = sheet.createRow(RowNum);
                    RowNum++;
                    for (int i = 0; i < str.length; i++) {
                        currentRow.createCell(i).setCellValue(str[i]);
                    }
                }catch (Exception e){
                    String message = e.getMessage();
                    System.out.println("message = " + message);
                }
            }
            FileOutputStream fileOutputStream = new FileOutputStream(excelFile);
            workBook.write(fileOutputStream);
            fileOutputStream.close();
            System.out.println("Done");
        } catch (Exception ex) {
            System.out.println(ex.getMessage() + "Exception in try");
        }
    }

    public static List listOfFiles() {
        List<String> results = new ArrayList<String>();
        File[] files = new File("C:\\MergeExcel").listFiles();

        for (File file : files) {
            if (file.isFile()) {
                results.add(file.getAbsolutePath());
            }
        }
        return results;
    }

    public static void main(String[] args) {
        List list = listOfFiles();
        int size = list.size();
        XSSFWorkbook workBook = new XSSFWorkbook();
        for (int i = 0; i < size; i++) {
            csvToXLSX(workBook,list.get(i).toString(), "sheet" + i + 1);
        }

        System.out.println("operation completed !!");
    }
}
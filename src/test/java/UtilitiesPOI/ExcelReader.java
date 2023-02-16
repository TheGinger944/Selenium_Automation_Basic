package UtilitiesPOI;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {
    public static Workbook wrkbook;
    public static Sheet wrksheet;

    public ExcelReader(String ExcelPath) throws IOException {
        try {
            File excel = new File(ExcelPath);
            FileInputStream file = new FileInputStream(excel);
            wrkbook = new XSSFWorkbook(file);
            wrksheet = wrkbook.getSheet("Sheet1");
        } catch (IOException e) {
            throw new IOException();
        }
    }

    public Cell getCell(int Column, int Row) {
        return wrksheet.getRow(Row).getCell(Column);
    }

    public void setValueIntoCell(int Column, int Row, String Value) {
        wrksheet.getRow(Row).getCell(Column).setCellValue(Value);
    }

    public void closeFile(String ExcelPath) throws IOException {
        try {
            FileOutputStream fos = new FileOutputStream(ExcelPath);
            wrkbook.write(fos);
            wrkbook.close();
        } catch (IOException e) {
            throw new IOException();
        }
    }
}
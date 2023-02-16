package UtilitiesJXL;

import java.io.File;
import java.io.IOException;
import java.util.Hashtable;

import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.WritableWorkbook;
import jxl.write.WritableSheet;

public class ExcelReader {
    public static Workbook wrkbook;
    public static WritableWorkbook wwbCopy;
    public static WritableSheet wrksheet;
    public static Hashtable<String, Integer> dict = new Hashtable<String, Integer>();

    public ExcelReader(String ExcelSheetPath) throws BiffException, IOException {
        try {
            wrkbook = Workbook.getWorkbook(new File(ExcelSheetPath));
            wwbCopy = Workbook.createWorkbook(new File(ExcelSheetPath), wrkbook);
            wrksheet = wwbCopy.getSheet("Sheet1");
        } catch (IOException e) {
            throw new IOException();
        }
    }

    public static int RowCount() {
        return wrksheet.getRows();
    }

    public static String ReadCell(int column, int row) {
        return wrksheet.getCell(column, row).getContents();
    }

    public static void ColumnDictionary() {

        for (int col = 0; col < wrksheet.getColumns(); col++) {
            dict.put(ReadCell(col, 0), col);
        }
    }

    public static int GetCell(String colName) {
        try {
            int value;
            value = ((Integer) dict.get(colName)).intValue();
            return value;
        } catch (NullPointerException e) {
            return (0);
        }
    }
}

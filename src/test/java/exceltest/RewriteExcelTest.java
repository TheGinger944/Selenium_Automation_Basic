package exceltest;

import jxl.read.biff.BiffException;
import jxl.write.WriteException;
import static org.junit.Assert.*;
import java.io.IOException;

public class RewriteExcelTest {
    public static void main(String[] args) throws BiffException, IOException, WriteException {
        // Create Objects
        ExcelReader excelReaderObj;
        CommonMethods commonMethodobj = new CommonMethods();
        TestData data = new TestData();
        //excelReaderObj = new ExcelReader("E:/Task3Project/excel/TestDataExcel.xls", "E:/Task3Project/excel/TestCopyExcel.xls");
        excelReaderObj = new ExcelReader("E:/Task3Project/excel/TestDataExcel.xls");
        excelReaderObj.ColumnDictionary();
        commonMethodobj.readExcelData(data);
        String origValue1 = data.getElement2().get(0);
        String origValue2 = data.getElement2().get(1);
        System.out.println(data.getElement2().get(0));
        System.out.println(data.getElement2().get(1));
        data.setValueIntoCell(4,1,"hohoho");
        data.setValueIntoCell(4,2,"banana");
        commonMethodobj.readExcelData(data);
        System.out.println(data.getElement2().get(0));
        System.out.println(data.getElement2().get(1));
        assertTrue(data.getElement2().get(1) != origValue1);
        assertTrue(data.getElement2().get(2) != origValue2);
        data.closeFile();
    }
}

package UtilitiesJXL.test;

import UtilitiesJXL.CommonMethods;
import UtilitiesJXL.ExcelReader;
import UtilitiesJXL.dataSetters.TestData;
import jxl.read.biff.BiffException;
import jxl.write.WriteException;

import static org.junit.Assert.*;

import java.io.IOException;

public class RewriteExcelTest {
    public static void main(String[] args) throws BiffException, IOException, WriteException {
        ExcelReader excelReaderObj;
        CommonMethods commonMethodobj = new CommonMethods();
        TestData data = new TestData();
        excelReaderObj = new ExcelReader("E:\\GitHubSeleniumRepo\\Selenium_Automation_Basic\\excel\\TestDataExcel.xls");
        excelReaderObj.ColumnDictionary();
        commonMethodobj.readExcelData(data);
        String origValue1 = data.getElement2().get(0);
        String origValue2 = data.getElement2().get(1);
        commonMethodobj.setValueIntoCell(4, 1, "hohoho");
        commonMethodobj.setValueIntoCell(4, 2, "banana");
        commonMethodobj.readExcelData(data);
        assertTrue(data.getElement2().get(1) != origValue1);
        assertTrue(data.getElement2().get(2) != origValue2);
        commonMethodobj.closeFile();
    }
}

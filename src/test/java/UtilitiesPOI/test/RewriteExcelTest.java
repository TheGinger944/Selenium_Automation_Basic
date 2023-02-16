package UtilitiesPOI.test;

import UtilitiesPOI.ExcelReader;
import static org.junit.Assert.*;
import java.io.IOException;

public class RewriteExcelTest {
    public static void main(String[] args) throws IOException {
        ExcelReader data;
        data = new ExcelReader("E:/GitHubSeleniumRepo/Selenium_Automation_Basic/excel/TestDataPOI.xlsx");
        String origValue1 = String.valueOf(data.getCell(4,1));
        String origValue2 = String.valueOf(data.getCell(4,2));
        data.setValueIntoCell(4,1,"hohoho");
        data.setValueIntoCell(4,2,"banana");
        assertTrue(String.valueOf(data.getCell(4,1)) != origValue1);
        assertTrue(String.valueOf(data.getCell(4,2)) != origValue2);
        data.closeFile("E:/GitHubSeleniumRepo/Selenium_Automation_Basic/excel/TestDataPOI.xlsx");
    }
}

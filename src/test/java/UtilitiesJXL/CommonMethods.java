package UtilitiesJXL;

import UtilitiesJXL.dataSetters.TestData;
import jxl.write.Label;
import jxl.write.WriteException;

import java.util.ArrayList;

import static UtilitiesJXL.ExcelReader.wrksheet;
import static UtilitiesJXL.ExcelReader.wwbCopy;
import static UtilitiesJXL.ExcelReader.wrkbook;

public class CommonMethods {
    public void readExcelData(TestData data) {
        ArrayList<String> browser = new ArrayList<String>();
        ArrayList<String> username = new ArrayList<String>();
        ArrayList<String> password = new ArrayList<String>();
        ArrayList<String> element1 = new ArrayList<String>();
        ArrayList<String> element2 = new ArrayList<String>();
        ArrayList<String> element3 = new ArrayList<String>();

        // Get the data from excel file
        for (int rowCnt = 1; rowCnt < ExcelReader.RowCount(); rowCnt++) {
            browser.add(ExcelReader.ReadCell(ExcelReader.GetCell("Browser"), rowCnt));
            username.add(ExcelReader.ReadCell(ExcelReader.GetCell("User ID"), rowCnt));
            password.add(ExcelReader.ReadCell(ExcelReader.GetCell("Password"), rowCnt));
            element1.add(ExcelReader.ReadCell(ExcelReader.GetCell("Element1"), rowCnt));
            element2.add(ExcelReader.ReadCell(ExcelReader.GetCell("Element2"), rowCnt));
            element3.add(ExcelReader.ReadCell(ExcelReader.GetCell("Element3"), rowCnt));
        }
        data.setBrowser(browser);
        data.setLoginUser(username);
        data.setPassword(password);
        data.setElement1(element1);
        data.setElement2(element2);
        data.setElement3(element3);
    }

    public void setValueIntoCell(int ColumnNumber, int RowNumber, String Value) throws WriteException {
        Label label = new Label(ColumnNumber, RowNumber, Value);
        try {
            wrksheet.addCell(label);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void closeFile() {
        try {
            wwbCopy.write();
            wwbCopy.close();
            wrkbook.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}

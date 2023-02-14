package exceltest;

import jxl.write.Label;
import jxl.write.WriteException;

import java.util.ArrayList;
import static exceltest.ExcelReader.wrkbook;
import static exceltest.ExcelReader.wrksheet;
import static exceltest.ExcelReader.wwbCopy;

public class TestData {
    private ArrayList<String> loginUser = null;
    private ArrayList<String> password = null;
    private ArrayList<String> browser = null;
    private ArrayList<String> element1 = null;
    private ArrayList<String> element2 = null;
    private ArrayList<String> element3 = null;

    public ArrayList<String> getLoginUser() {
        return loginUser;
    }

    public void setLoginUser(ArrayList<String> loginUser) {
        this.loginUser = loginUser;
    }

    public ArrayList<String> getPassword() {
        return password;
    }

    public void setPassword(ArrayList<String> password) {
        this.password = password;
    }

    public ArrayList<String> getBrowser() {
        return browser;
    }

    public void setBrowser(ArrayList<String> browser) {
        this.browser = browser;
    }

    public ArrayList<String> getElement1() {
        return element1;
    }

    public void setElement1(ArrayList<String> element1) {
        this.element1 = element1;
    }

    public ArrayList<String> getElement2() {
        return element2;
    }

    public void setElement2(ArrayList<String> element2) {
        this.element2 = element2;
    }

    public ArrayList<String> getElement3() {
        return element3;
    }

    public void setElement3(ArrayList<String> element3) {
        this.element3 = element3;
    }

    public void setValueIntoCell(int ColumnNumber, int RowNumber,String Value) throws WriteException
    {
        Label label = new Label(ColumnNumber, RowNumber, Value);
        try {
            wrksheet.addCell(label);
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }

    public void closeFile()
    {
        try {
            // Closing the writable work book
            wwbCopy.write();
            wwbCopy.close();
            wrkbook.close();
        } catch (Exception e)

        {
            e.printStackTrace();
        }
    }
}

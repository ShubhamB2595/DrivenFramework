package test;

import java.io.IOException;
import java.util.Properties;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import operations.ObjectOperation;
import operations.ReadObject;
import utility.ExcelUtil;

public class TestCaseHybrid {

	WebDriver driver;

	@Test(dataProvider = "hybridData")
	public void testLogin(String testcaseName, String keyword, String objectName, String objectType, String value)
			throws Exception {
		// TODO Auto-generated method stub

		if (testcaseName != null && testcaseName.length() != 0) {
			driver = new ChromeDriver();
		}

		ReadObject object = new ReadObject();
		Properties allObjects = object.getObjectRepository();
		ObjectOperation operation = new ObjectOperation(driver);

		// Call perform function to perform operation on UI
		operation.perform(allObjects, keyword, objectName, objectType, value);

	}

	@DataProvider(name = "hybridData")
	public Object[][] getDataFromDataprovider() throws IOException {

		Object[][] object = null;
		ExcelUtil file = new ExcelUtil();

		// Read keyword sheet
		Sheet mySheet = file.readExcel(System.getProperty("user.dir") + "\\", "TestCase.xlsx", "KeywordFramework");

		// Find number of rows in excel file
		int rowCount = mySheet.getLastRowNum() - mySheet.getFirstRowNum();
		object = new Object[rowCount][5];

		for (int i = 0; i < rowCount; i++) {
			// Loop over all the rows
			Row row = mySheet.getRow(i + 1);
			// Create a loop to print cell values in a row
			for (int j = 0; j < row.getLastCellNum(); j++) {
				// Print excel data in console
				object[i][j] = row.getCell(j).toString();
			}
		}

		System.out.println("");
		return object;
	}
	
	@AfterTest
	public void close() {
		
		driver.close();
	}
}

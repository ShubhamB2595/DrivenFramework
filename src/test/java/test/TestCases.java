package test;

import java.util.Properties;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.Test;

import operations.ObjectOperation;
import operations.ReadObject;
import utility.ExcelUtil;

public class TestCases {

	@Test
	public void signInTest() throws Exception {

		System.setProperty("webdriver.chrome.driver", "I:\\Selenium\\KeywordsDriven\\chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		
		ExcelUtil file = new ExcelUtil();
		ReadObject object = new ReadObject();
		Properties allObjects = object.getObjectRepository();
		ObjectOperation operation = new ObjectOperation(driver);

		// Read keyword sheet
		Sheet mySheet = file.readExcel(System.getProperty("user.dir") + "\\", "TestCase.xlsx", "KeywordFramework");

		// Find number of rows in excel file
		int rowCount = mySheet.getLastRowNum() - mySheet.getFirstRowNum();

		// Create a loop over all the rows of excel file to read it
		for (int i = 1; i < rowCount + 1; i++) {

			// Loop over all the rows
			Row row = mySheet.getRow(i);

			// Check if the first cell contain a value, if yes, That means it is the new
			// testcase name
			if (row.getCell(0).toString().length() == 0) {

				// Print testcase detail on console
				System.out.println(row.getCell(1).toString() + "----" + row.getCell(2).toString() + "----"
						+ row.getCell(3).toString() + "----" + row.getCell(4).toString());

				// Call perform function to perform operation on UI
				operation.perform(allObjects, row.getCell(1).toString(), row.getCell(2).toString(),
						row.getCell(3).toString(), row.getCell(4).toString());
			} else {

				// Print the new testcase name when it started
				System.out.println("New Testcase->" + row.getCell(0).toString() + " Started");
			}
		}
	}
}

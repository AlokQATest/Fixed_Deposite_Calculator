package Data_Driven_Testing_FD_Calculator;

import org.testng.annotations.Test;
import org.testng.annotations.Test;
import java.io.IOException;
import java.time.Duration;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.Test;

public class FD_Calculator {
	
	WebDriver driver;
	static String file_Path = "/Users/aloktiwari/eclipse-workspace/SeleniumJava/testData/Test_Data.xlsx";
	static String sheet_Name = "Sheet1";
	
	
	@Test(priority=1)
	void open_App() {
	driver = new ChromeDriver();
    driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));
    driver.get("https://www.moneycontrol.com/fixed-income/calculator/state-bank-of-india-sbi/fixed-deposit-calculator-SBI-BSB001.html?classic=true");
    driver.manage().window().maximize();
	}
	@Test(priority=2)
	int rowCount() throws IOException {
		return excelUtils.getRowCount(file_Path, sheet_Name);
	}
	@Test(priority=3)
	int cellCount() throws IOException {
		return excelUtils.getCellCount(file_Path, sheet_Name);
	}
	
	@Test(priority=4)
	void mainFunction() throws NumberFormatException, IOException {
		for (int r = 1; r <= rowCount(); r++) {

            //Reading cell data from Excel

            String principal = excelUtils.getCellData(file_Path, sheet_Name, r, 0);
            String rate = excelUtils.getCellData(file_Path, sheet_Name, r, 1);
            String period = excelUtils.getCellData(file_Path, sheet_Name, r, 2);
            String tenurePeriod = excelUtils.getCellData(file_Path, sheet_Name, r, 3);
            String frequency = excelUtils.getCellData(file_Path, sheet_Name, r, 4);
            String maturityValue = excelUtils.getCellData(file_Path, sheet_Name, r, 5);

            // Passing above data into web application
            
            driver.findElement(By.xpath("//input[@name='principal']")).sendKeys(principal);
            driver.findElement(By.xpath("//input[@name='interest']")).sendKeys(rate);
            driver.findElement(By.xpath("//input[@name='tenure']")).sendKeys(period);

            Select ten = new Select(driver.findElement(By.xpath("//select[@name='tenurePeriod']")));
            ten.selectByVisibleText(tenurePeriod);

            Select freq = new Select(driver.findElement(By.xpath("//select[@name= 'frequency']")));
            freq.selectByVisibleText(frequency);


            driver.findElement(By.xpath("//img[@src='https://images.moneycontrol.com/images/mf_revamp/btn_calcutate.gif']")).click();

            //Data validation
            
            String matValue = driver.findElement(By.xpath("//span[@id ='resp_matval']//strong")).getText();
        
            if (Double.parseDouble(maturityValue) == Double.parseDouble(matValue)) {
                excelUtils.setCellData(file_Path, sheet_Name, r, 7, "Passed");
                excelUtils.setGreenColor(file_Path, sheet_Name, r, 7);
            } else {
                excelUtils.setCellData(file_Path, sheet_Name, r, 7, "Failed");
                excelUtils.setRedColor(file_Path, sheet_Name, r, 7);
            }


            driver.findElement(By.xpath("//img[@src='https://images.moneycontrol.com/images/mf_revamp/btn_clear.gif']")).click();
        }
	}
	
	
}

package Data_Driven_Testing_FD_Calculator;

import java.io.IOException;
import java.time.Duration;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

public class FD_Calculator_Main {

	public static void main(String[] args) throws IOException {
		
		WebDriver driver = new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));
		driver.get("https://www.moneycontrol.com/fixed-income/calculator/state-bank-of-india-sbi/fixed-deposit-calculator-SBI-BSB001.html?classic=true");
		driver.manage().window().maximize();
		
		
		
		
		String file = "/Users/aloktiwari/eclipse-workspace/SeleniumJava/testData/Test_Data.xlsx";
		int row_Count = excelUtills.getRowCount(file, "Sheet1");
		int cell_Count = excelUtills.getCellCount(file, "Sheet1");
		
		for(int r=1; r<=row_Count; r++) {
			
			//Reading cell data from Excel
				
			String principal = excelUtills.getCellData(file, "Sheet1", r, 0);
			String rate = excelUtills.getCellData(file, "Sheet1", r,1);
			String period = excelUtills.getCellData(file, "Sheet1", r, 2);
			String tenurePeriod = excelUtills.getCellData(file, "Sheet1", r, 3);
			String frequency = excelUtills.getCellData(file, "Sheet1", r, 4);
			String maturityValue = excelUtills.getCellData(file, "Sheet1", r, 5);
			System.out.println(maturityValue);
			
			
			// Passing above data into web application
			driver.findElement(By.xpath("//input[@name='principal']")).sendKeys(principal);
			driver.findElement(By.xpath("//input[@name='interest']")).sendKeys(rate);
			driver.findElement(By.xpath("//input[@name='tenure']")).sendKeys(period);
			//driver.findElement(By.xpath("//select[@name='tenurePeriod']")).click();
			
			Select ten = new Select(driver.findElement(By.xpath("//select[@name='tenurePeriod']")));
			ten.selectByVisibleText(tenurePeriod);
			
			Select freq = new Select(driver.findElement(By.xpath("//select[@name= 'frequency']")));
			freq.selectByVisibleText(frequency);
	
			
			driver.findElement(By.xpath("//img[@src='https://images.moneycontrol.com/images/mf_revamp/btn_calcutate.gif']")).click();
				
			//Data validation
			String matValue = driver.findElement(By.xpath("//span[@id ='resp_matval']//strong")).getText();
			System.out.println(matValue);
			if(Double.parseDouble(maturityValue)==Double.parseDouble(matValue)) {
				excelUtills.setCellData(file, "Sheet1", r, 7, "Passed");
				excelUtills.setGreenColor(file, "Sheet1", r, 7);
			}else {
			excelUtills.setCellData(file, "Sheet1", r, 7, "Failed");
			excelUtills.setRedColor(file, "Sheet1", r, 7);
			}
			
				
			driver.findElement(By.xpath("//img[@src='https://images.moneycontrol.com/images/mf_revamp/btn_clear.gif']")).click();
			}
		}
	}


